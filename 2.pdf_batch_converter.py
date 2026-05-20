'''
@Project ：PycharmProjects
@File    ：年报批量下载.py
@IDE     ：PyCharm
@Author  ：lingxiaotian
@Date    ：2023/5/30 11:39
@LastEditTime: 2025/11/21 14:10
'''

from __future__ import annotations

import logging
import os
import re
import warnings
import webbrowser
import wsgiref.simple_server
import wsgiref.util
from dataclasses import dataclass
from multiprocessing import Pool, cpu_count
from pathlib import Path
from typing import Any, Optional, Tuple

import pandas as pd
import pdfplumber
import requests

# 尝试导入备用PDF处理库
try:
    from PyPDF2 import PdfReader
    PYPDF2_AVAILABLE = True
except ImportError:
    PYPDF2_AVAILABLE = False
    logging.warning("PyPDF2 not installed. Install with: pip install PyPDF2")

try:
    from pdfminer.high_level import extract_text as pdfminer_extract
    PDFMINER_AVAILABLE = True
except ImportError:
    PDFMINER_AVAILABLE = False
    logging.warning("pdfminer.six not installed. Install with: pip install pdfminer.six")

# Google Drive API (OAuth)
try:
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload
    GDRIVE_AVAILABLE = True
except ImportError:
    GDRIVE_AVAILABLE = False
    logging.warning(
        "Google Drive libraries not installed. "
        "Install with: pip install google-api-python-client google-auth-oauthlib google-auth-httplib2"
    )

# 抑制pdfplumber的CropBox警告
warnings.filterwarnings('ignore', message='.*CropBox.*')

# 日志配置
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)


@dataclass(frozen=True)
class ConverterConfig:
    """PDF批量下载转换配置类。"""
    excel_file: str  # Excel表格路径
    pdf_dir: str  # PDF存储目录
    txt_dir: str  # TXT存储目录
    target_year: int  # 目标年份
    delete_pdf: bool = False  # 是否删除转换后的PDF
    max_retries: int = 3  # 下载最大重试次数
    timeout: int = 15  # 请求超时时间（秒）
    chunk_size: int = 8192  # 下载块大小
    processes: Optional[int] = None  # 进程数，None表示自动
    # Google Drive (OAuth)：仅当 upload_to_gdrive=True 时生效
    upload_to_gdrive: bool = False
    delete_local_after_upload: bool = True  # 上传成功后删除本地 PDF 与 TXT
    gdrive_folder_id: str = ""
    gdrive_credentials_path: str = ""
    gdrive_token_path: str = "token.json"
    upload_record_file: str = "uploaded_txt_records.txt"  # 本地上传完成记录（每行一个 txt 文件名）


# Google Drive OAuth scope（仅访问通过本应用创建/打开的文件）
GDRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.file"]

# 固定端口避免 redirect_uri 与 Console 不一致导致 400；在 Console 已授权重定向 URI 中添加 http://localhost:8080/
_GDRIVE_OAUTH_PORT = 8080
_DEBUG_LOG_PATH = "/Users/evawang/Desktop/RA-ZY/Annualreport_tools/.cursor/debug.log"


class _OAuthRedirectWSGIApp:
    """WSGI app 记录回调 URI，不修改 http/https。"""
    def __init__(self, success_message: str):
        self.success_message = success_message
        self.last_request_uri = None

    def __call__(self, environ, start_response):
        start_response("200 OK", [("Content-type", "text/plain; charset=utf-8")])
        self.last_request_uri = wsgiref.util.request_uri(environ)
        return [self.success_message.encode("utf-8")]


def _run_local_server_avoid_https_redirect(flow: Any, port: int) -> Any:
    """与 run_local_server 相同，但 fetch_token 时不再把 http 改成 https，避免 redirect_uri 不一致导致 400。"""
    success_message = "认证完成，可关闭此页面。"
    wsgi_app = _OAuthRedirectWSGIApp(success_message)
    local_server = wsgiref.simple_server.make_server("localhost", port, wsgi_app)
    flow.redirect_uri = "http://localhost:{}/".format(local_server.server_port)
    auth_url, _ = flow.authorization_url(prompt="consent", access_type="offline")
    # #region agent log
    _debug_log("auth URL generated", {"url_preview": auth_url[:350] if auth_url else None, "redirect_uri": flow.redirect_uri}, "H5")
    # #endregion
    webbrowser.open(auth_url, new=1, autoraise=True)
    logging.info("请在浏览器中完成授权，若看到 400 请查看上方日志中的 auth URL 与 Console 配置是否一致")
    local_server.handle_request()
    if not wsgi_app.last_request_uri:
        raise RuntimeError("未收到授权回调，请重试")
    # 不把 http 改为 https，保持与授权请求中的 redirect_uri 一致。
    # oauthlib 默认拒绝 http 回调，这里仅对 localhost 授权流程临时放开。
    old_insecure_transport = os.environ.get("OAUTHLIB_INSECURE_TRANSPORT")
    os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"
    # #region agent log
    _debug_log(
        "setting OAUTHLIB_INSECURE_TRANSPORT",
        {"value": "1", "old_value": old_insecure_transport},
        "H6",
    )
    # #endregion
    try:
        flow.fetch_token(authorization_response=wsgi_app.last_request_uri)
    finally:
        if old_insecure_transport is None:
            os.environ.pop("OAUTHLIB_INSECURE_TRANSPORT", None)
        else:
            os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = old_insecure_transport
    return flow.credentials


def _debug_log(message: str, data: dict, hypothesis_id: str) -> None:
    import json
    try:
        with open(_DEBUG_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(json.dumps({"message": message, "data": data, "hypothesisId": hypothesis_id, "timestamp": __import__("time").time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass


def _get_gdrive_service(credentials_path: str, token_path: str) -> Any:
    """使用 OAuth 构建 Drive API 服务。无 token 时打开浏览器完成授权并保存 token。"""
    # #region agent log
    _debug_log("_get_gdrive_service entry", {"credentials_path": credentials_path, "token_path": token_path, "token_exists": os.path.exists(token_path), "gdrive_available": GDRIVE_AVAILABLE}, "H1")
    # #endregion
    if not GDRIVE_AVAILABLE:
        raise RuntimeError("Google Drive 库未安装，无法上传")
    creds = None
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, GDRIVE_SCOPES)
    # #region agent log
    _debug_log("after load token", {"creds_is_none": creds is None, "creds_valid": getattr(creds, "valid", None) if creds else None}, "H3")
    # #endregion
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(credentials_path):
                raise FileNotFoundError(f"OAuth 客户端密钥不存在: {credentials_path}")
            # #region agent log
            _debug_log("starting OAuth flow", {"redirect_uri_will_be": f"http://localhost:{_GDRIVE_OAUTH_PORT}/", "hint": "Add this exact URI in Google Cloud Console > Credentials > OAuth client > Authorized redirect URIs"}, "H5")
            # #endregion
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, GDRIVE_SCOPES)
            try:
                creds = _run_local_server_avoid_https_redirect(flow, _GDRIVE_OAUTH_PORT)
            except Exception as e:
                # #region agent log
                _debug_log("run_local_server exception", {"type": type(e).__name__, "message": str(e)}, "H2")
                # #endregion
                raise
            # #region agent log
            _debug_log("OAuth flow completed", {"creds_got": creds is not None}, "H4")
            # #endregion
        with open(token_path, "w") as f:
            f.write(creds.to_json())
    return build("drive", "v3", credentials=creds)


def _upload_file_to_gdrive(
    service: Any,
    local_path: str,
    parent_folder_id: str,
    drive_name: Optional[str] = None,
) -> bool:
    """上传本地文件到 Drive 指定文件夹。成功返回 True。"""
    if not os.path.exists(local_path):
        logging.error(f"上传失败，文件不存在: {local_path}")
        return False
    name = drive_name or os.path.basename(local_path)
    try:
        media = MediaFileUpload(local_path, resumable=True)
        service.files().create(
            body={"name": name, "parents": [parent_folder_id]},
            media_body=media,
            fields="id",
        ).execute()
        logging.info(f"已上传到 Drive: {name}")
        return True
    except Exception as e:
        logging.error(f"上传到 Drive 失败 {local_path}: {e}")
        return False


def _gdrive_txt_exists(service: Any, parent_folder_id: str, drive_name: str) -> bool:
    """检查目标 Drive 文件夹中是否已存在同名文件。"""
    safe_name = drive_name.replace("'", "\\'")
    query = (
        f"name = '{safe_name}' and "
        f"'{parent_folder_id}' in parents and "
        "trashed = false"
    )
    try:
        resp = service.files().list(
            q=query,
            fields="files(id,name)",
            pageSize=1,
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
        ).execute()
        files = resp.get("files", [])
        return len(files) > 0
    except Exception as e:
        logging.warning(f"查询 Drive 同名文件失败，将继续尝试上传: {e}")
        return False


class PDFDownloader:
    """PDF下载器类。"""
    
    HEADERS = {
        "Accept-Encoding": "gzip, deflate, br",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        "Cache-Control": "no-cache",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "same-origin",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    }
    
    def __init__(self, timeout: int = 15, chunk_size: int = 8192) -> None:
        self.timeout = timeout
        self.chunk_size = chunk_size
        self.session = requests.Session()
        self.session.headers.update(self.HEADERS)
    
    def download(self, pdf_url: str, pdf_file_path: str) -> bool:
        """下载PDF文件并验证完整性。
        
        Args:
            pdf_url: PDF下载链接
            pdf_file_path: 保存路径
            
        Returns:
            下载是否成功
        """
        try:
            # 请求PDF文件
            response = self.session.get(pdf_url, stream=True, timeout=self.timeout)
            
            # 检查HTTP状态码
            if response.status_code == 403:
                logging.error(f"403 Forbidden: 服务器禁止访问 {pdf_url}")
                return False
            elif response.status_code != 200:
                logging.error(f"请求失败: {response.status_code} - {response.text[:500]}")
                return False
            
            # 验证Content-Type
            content_type = response.headers.get("Content-Type", "")
            if "pdf" not in content_type.lower():
                logging.error(f"服务器返回的不是 PDF: {content_type}")
                return False
            
            # 写入PDF文件
            with open(pdf_file_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=self.chunk_size):
                    if chunk:
                        f.write(chunk)
            
            # 验证文件完整性
            if not self._verify_pdf(pdf_file_path):
                return False
            
            logging.info(f"PDF 下载成功: {pdf_file_path}")
            return True
            
        except requests.exceptions.Timeout:
            logging.error(f"下载超时: {pdf_url}")
            return False
        except requests.exceptions.RequestException as e:
            logging.error(f"下载 PDF 文件失败: {e}")
            return False
        except OSError as e:
            logging.error(f"文件写入失败: {e}")
            return False
    
    @staticmethod
    def _verify_pdf(pdf_file_path: str) -> bool:
        """验证PDF文件完整性。"""
        if not os.path.exists(pdf_file_path):
            logging.error(f"文件不存在: {pdf_file_path}")
            return False
        
        if os.path.getsize(pdf_file_path) == 0:
            logging.error(f"下载失败，文件大小为 0 KB: {pdf_file_path}")
            return False
        
        try:
            with open(pdf_file_path, "rb") as f:
                first_bytes = f.read(5)
                if not first_bytes.startswith(b"%PDF"):
                    logging.error(f"下载的文件不是有效的 PDF: {pdf_file_path}")
                    return False
        except OSError as e:
            logging.error(f"文件验证失败: {e}")
            return False
        
        return True

class PDFConverter:
    """PDF转TXT转换器类。"""
    
    # 文件名非法字符正则
    INVALID_CHARS = r'[\\/:*?"<>|]'
    
    def __init__(self, config: ConverterConfig) -> None:
        self.config = config
        self.downloader = PDFDownloader(
            timeout=config.timeout,
            chunk_size=config.chunk_size
        )
        self._drive_service: Any = None
        self._gdrive_folder_id: str = ""
        self._uploaded_records: set[str] = set()
        self._upload_record_path = Path(self.config.upload_record_file)
        self._load_uploaded_records()
    
    def set_drive_service(self, service: Any, folder_id: str) -> None:
        """注入 Drive API 服务与目标文件夹 ID（上传到 Drive 时由 Processor 调用）。"""
        self._drive_service = service
        self._gdrive_folder_id = folder_id

    def _load_uploaded_records(self) -> None:
        """读取本地上传记录文件。"""
        if not self._upload_record_path.exists():
            return
        try:
            with open(self._upload_record_path, "r", encoding="utf-8") as f:
                self._uploaded_records = {line.strip() for line in f if line.strip()}
            logging.info(f"已加载上传记录: {self._upload_record_path} ({len(self._uploaded_records)} 条)")
        except OSError as e:
            logging.warning(f"读取上传记录失败: {e}")

    def _mark_uploaded_record(self, txt_name: str) -> None:
        """写入本地上传记录（去重）。"""
        if txt_name in self._uploaded_records:
            return
        self._uploaded_records.add(txt_name)
        try:
            self._upload_record_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self._upload_record_path, "a", encoding="utf-8") as f:
                f.write(txt_name + "\n")
        except OSError as e:
            logging.warning(f"写入上传记录失败: {e}")
    
    @staticmethod
    def _sanitize_filename(filename: str) -> str:
        """清理文件名中的非法字符。"""
        return re.sub(PDFConverter.INVALID_CHARS, '', filename)
    
    def _download_with_retry(self, pdf_url: str, pdf_file_path: str) -> bool:
        """带重试机制的下载。"""
        for attempt in range(1, self.config.max_retries + 1):
            if self.downloader.download(pdf_url, pdf_file_path):
                return True
            if attempt < self.config.max_retries:
                logging.warning(f"重试下载 ({attempt}/{self.config.max_retries}): {pdf_url}")
        
        logging.error(f"下载失败（已重试 {self.config.max_retries} 次）: {pdf_url}")
        return False
    
    def _convert_with_pdfplumber(self, pdf_path: str, txt_path: str) -> bool:
        """使用pdfplumber转换PDF（方法1）。"""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                with open(txt_path, 'w', encoding='utf-8') as f:
                    for page_num, page in enumerate(pdf.pages, 1):
                        try:
                            text = page.extract_text()
                            if text:
                                f.write(text)
                        except Exception as e:
                            logging.debug(f"pdfplumber提取第 {page_num} 页失败: {e}")
                            continue
            return True
        except Exception as e:
            logging.debug(f"pdfplumber转换失败: {e}")
            return False
    
    def _convert_with_pypdf2(self, pdf_path: str, txt_path: str) -> bool:
        """使用PyPDF2转换PDF（方法2）。"""
        if not PYPDF2_AVAILABLE:
            return False
        
        try:
            reader = PdfReader(pdf_path)
            with open(txt_path, 'w', encoding='utf-8') as f:
                for page_num, page in enumerate(reader.pages, 1):
                    try:
                        text = page.extract_text()
                        if text:
                            f.write(text)
                    except Exception as e:
                        logging.debug(f"PyPDF2提取第 {page_num} 页失败: {e}")
                        continue
            return True
        except Exception as e:
            logging.debug(f"PyPDF2转换失败: {e}")
            return False
    
    def _convert_with_pdfminer(self, pdf_path: str, txt_path: str) -> bool:
        """使用pdfminer.six转换PDF（方法3）。"""
        if not PDFMINER_AVAILABLE:
            return False
        
        try:
            text = pdfminer_extract(pdf_path)
            with open(txt_path, 'w', encoding='utf-8') as f:
                f.write(text)
            return True
        except Exception as e:
            logging.debug(f"pdfminer转换失败: {e}")
            return False
    
    def _convert_pdf_to_txt(self, pdf_path: str, txt_path: str) -> bool:
        """将PDF转换为TXT，使用多种库作为备用方案。
        
        尝试顺序：
        1. pdfplumber（默认，最准确）
        2. PyPDF2（备用方案1）
        3. pdfminer.six（备用方案2）
        """
        # 方法1: pdfplumber
        if self._convert_with_pdfplumber(pdf_path, txt_path):
            logging.info(f"转换成功 (pdfplumber): {txt_path}")
            return True
        
        logging.warning(f"pdfplumber转换失败，尝试备用方案: {pdf_path}")
        
        # 方法2: PyPDF2
        if self._convert_with_pypdf2(pdf_path, txt_path):
            logging.info(f"转换成功 (PyPDF2): {txt_path}")
            return True
        
        # 方法3: pdfminer.six
        if self._convert_with_pdfminer(pdf_path, txt_path):
            logging.info(f"转换成功 (pdfminer): {txt_path}")
            return True
        
        logging.error(f"所有PDF转换方法均失败: {pdf_path}")
        return False
    
    def process_single_file(
        self,
        code: Any,
        name: str,
        year: Any,
        pdf_url: str
    ) -> bool:
        """处理单个文件的下载和转换；若启用则上传到 Drive 并可选删除本地文件。
        
        Args:
            code: 公司代码（int 或 Excel 读出的类型）
            name: 公司简称
            year: 年份
            pdf_url: PDF下载链接
            
        Returns:
            处理是否成功
        """
        code_int = int(code) if not isinstance(code, int) else code
        base_name = self._sanitize_filename(f"{code_int:06d}_{name}_{year}")
        pdf_file_path = os.path.join(self.config.pdf_dir, f"{base_name}.pdf")
        txt_file_path = os.path.join(self.config.txt_dir, f"{base_name}.txt")
        txt_name = os.path.basename(txt_file_path)
        
        try:
            # 已在本地上传记录中：直接跳过
            if self.config.upload_to_gdrive and txt_name in self._uploaded_records:
                logging.info(f"本地记录显示已上传，跳过: {txt_name}")
                return True

            # 检查TXT是否已存在（且未开启上传时，存在即跳过；开启上传时若本地无则需下载转换后上传）
            if os.path.exists(txt_file_path) and not self.config.upload_to_gdrive:
                logging.info(f"文件已存在，跳过: {base_name}.txt")
                return True
            
            # 下载PDF（如果不存在）
            if not os.path.exists(pdf_file_path):
                if not self._download_with_retry(pdf_url, pdf_file_path):
                    return False
            
            # 转换PDF为TXT
            if not self._convert_pdf_to_txt(pdf_file_path, txt_file_path):
                return False
            
            if self.config.upload_to_gdrive and self._drive_service and self._gdrive_folder_id:
                # 上传到 Drive：仅上传 TXT（不上传 PDF）
                drive_txt_name = txt_name
                if _gdrive_txt_exists(self._drive_service, self._gdrive_folder_id, drive_txt_name):
                    logging.info(f"Drive 已存在同名 TXT，跳过上传: {drive_txt_name}")
                    txt_ok = True
                else:
                    txt_ok = _upload_file_to_gdrive(
                        self._drive_service,
                        txt_file_path,
                        self._gdrive_folder_id,
                        drive_txt_name,
                    )
                if not txt_ok:
                    logging.warning(f"TXT 上传失败，保留本地文件: {base_name}")
                    return False
                self._mark_uploaded_record(drive_txt_name)
                if self.config.delete_local_after_upload:
                    for p in (txt_file_path, pdf_file_path):
                        if os.path.exists(p):
                            try:
                                os.remove(p)
                                logging.info(f"已删除本地: {p}")
                            except OSError as e:
                                logging.warning(f"删除本地文件失败: {e}")
            else:
                # 未启用上传：按原逻辑可选删除 PDF
                if self.config.delete_pdf and os.path.exists(pdf_file_path):
                    try:
                        os.remove(pdf_file_path)
                        logging.info(f"已删除PDF: {pdf_file_path}")
                    except OSError as e:
                        logging.warning(f"删除PDF失败: {e}")
            
            return True
            
        except Exception as e:
            logging.error(f"处理文件失败 {code_int:06d}_{name}_{year}: {e}")
            return False



def _process_task(args: Tuple) -> bool:
    """多进程任务包装函数。"""
    converter, code, name, year, pdf_url = args
    return converter.process_single_file(code, name, year, pdf_url)


class AnnualReportProcessor:
    """年报批量处理器。"""
    
    def __init__(self, config: ConverterConfig) -> None:
        self.config = config
        self.converter = PDFConverter(config)
    
    def _load_excel_data(self) -> Optional[pd.DataFrame]:
        """加载Excel数据。"""
        try:
            df = pd.read_excel(self.config.excel_file)
            logging.info(f"成功加载Excel文件: {self.config.excel_file}")
            return df
        except FileNotFoundError:
            logging.error(f"Excel文件不存在: {self.config.excel_file}")
            return None
        except Exception as e:
            logging.error(f"读取Excel失败: {e}")
            return None
    
    def _prepare_directories(self) -> bool:
        """创建必要的目录。"""
        try:
            Path(self.config.pdf_dir).mkdir(parents=True, exist_ok=True)
            Path(self.config.txt_dir).mkdir(parents=True, exist_ok=True)
            logging.info(f"目录准备完成: PDF={self.config.pdf_dir}, TXT={self.config.txt_dir}")
            return True
        except OSError as e:
            logging.error(f"创建目录失败: {e}")
            return False
    
    def _filter_data_by_year(self, df: pd.DataFrame) -> pd.DataFrame:
        """按年份过滤数据。"""
        required_columns = ['公司代码', '公司简称', '年份', '年报链接']
        
        # 检查必需列
        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            logging.error(f"Excel缺少必需列: {missing_cols}")
            return pd.DataFrame()
        
        # 过滤年份
        filtered = df[df['年份'].astype(str) == str(self.config.target_year)]
        logging.info(f"找到 {len(filtered)} 条 {self.config.target_year} 年的记录")
        return filtered
    
    def run(self) -> None:
        """执行批量处理流程。启用上传时单进程顺序执行以避免 OAuth token 并发。"""
        logging.info("="*60)
        logging.info("年报批量下载转换程序启动")
        logging.info(f"目标年份: {self.config.target_year}")
        logging.info(f"删除PDF: {self.config.delete_pdf}")
        logging.info(f"上传到 Google Drive: {self.config.upload_to_gdrive}")
        if self.config.upload_to_gdrive:
            logging.info(f"上传后删除本地: {self.config.delete_local_after_upload}")
        logging.info("="*60)
        
        # 加载数据
        df = self._load_excel_data()
        if df is None:
            return
        
        # 准备目录
        if not self._prepare_directories():
            return
        
        # 过滤数据
        filtered_df = self._filter_data_by_year(df)
        if filtered_df.empty:
            logging.warning(f"未找到 {self.config.target_year} 年的数据")
            return
        
        tasks = [
            (self.converter, row['公司代码'], row['公司简称'], row['年份'], row['年报链接'])
            for _, row in filtered_df.iterrows()
        ]
        
        if self.config.upload_to_gdrive:
            if not self.config.gdrive_folder_id or not self.config.gdrive_credentials_path:
                logging.error("上传到 Drive 已开启，但未配置 gdrive_folder_id 或 gdrive_credentials_path")
                return
            try:
                service = _get_gdrive_service(
                    self.config.gdrive_credentials_path,
                    self.config.gdrive_token_path,
                )
                self.converter.set_drive_service(service, self.config.gdrive_folder_id)
            except Exception as e:
                logging.error(f"初始化 Google Drive 失败: {e}")
                return
            logging.info("使用单进程顺序处理（上传到 Drive）")
            success_count = 0
            for args in tasks:
                if _process_task(args):
                    success_count += 1
        else:
            worker_count = self.config.processes or min(cpu_count(), len(tasks))
            logging.info(f"使用 {worker_count} 个进程处理 {len(tasks)} 个文件")
            success_count = 0
            with Pool(processes=worker_count) as pool:
                results = pool.map(_process_task, tasks)
                success_count = sum(results)
        
        logging.info("="*60)
        logging.info(f"处理完成: 成功 {success_count}/{len(tasks)}")
        logging.info("="*60)


if __name__ == '__main__':
    # ==================== 配置区域 ====================
    
    # Excel表格路径（建议使用绝对路径）
    # 2024年02月14日更新后，此处只需要填写总表的路径，请于网盘或github中获取总表
    EXCEL_FILE = "年报链接_2024【公众号：凌小添】.xlsx"
    
    # 是否删除转换后的PDF文件（节省磁盘空间）
    DELETE_PDF = False
    
    # 是否批量处理多个年份
    BATCH_MODE = True
    
    # 批量模式：年份区间（包含起始和结束年份）
    START_YEAR = 2022
    END_YEAR = 2024
    
    # 单独模式：指定年份
    SINGLE_YEAR = 2023
    
    # 下载配置
    MAX_RETRIES = 3  # 最大重试次数
    TIMEOUT = 15  # 请求超时（秒）
    PROCESSES = None  # 进程数（None表示自动）
    
    # Google Drive（OAuth）：上传到共享文件夹，上传成功后删除本地以省空间
    UPLOAD_TO_GDRIVE = True
    DELETE_LOCAL_AFTER_UPLOAD = True
    GDRIVE_FOLDER_ID = "1L2qkRqXhSmhTMc7pmPYICSnR2EOhsBA_"  # 目标文件夹 ID（Drive 地址栏 folders/ 后的那串）
    GDRIVE_CREDENTIALS_PATH = "credentials.json"  # OAuth 客户端密钥 JSON
    GDRIVE_TOKEN_PATH = "token.json"  # 首次授权后保存的 token
    UPLOAD_RECORD_FILE = "uploaded_txt_records.txt"  # 本地上传记录（每行一个 txt 文件名）
    
    # ==================== 执行逻辑 ====================
    
    if BATCH_MODE:
        # 批量处理多个年份
        for year in range(START_YEAR, END_YEAR + 1):
            config = ConverterConfig(
                excel_file=EXCEL_FILE,
                pdf_dir=f'年报文件/{year}/pdf年报',
                txt_dir=f'年报文件/{year}/txt年报',
                target_year=year,
                delete_pdf=DELETE_PDF,
                max_retries=MAX_RETRIES,
                timeout=TIMEOUT,
                processes=PROCESSES,
                upload_to_gdrive=UPLOAD_TO_GDRIVE,
                delete_local_after_upload=DELETE_LOCAL_AFTER_UPLOAD,
                gdrive_folder_id=GDRIVE_FOLDER_ID,
                gdrive_credentials_path=GDRIVE_CREDENTIALS_PATH,
                gdrive_token_path=GDRIVE_TOKEN_PATH,
                upload_record_file=UPLOAD_RECORD_FILE,
            )
            
            processor = AnnualReportProcessor(config)
            processor.run()
            
            print(f"\n{year}年年报处理完毕\n")
    else:
        # 处理单独年份
        config = ConverterConfig(
            excel_file=EXCEL_FILE,
            pdf_dir=f'年报文件/{SINGLE_YEAR}/pdf年报',
            txt_dir=f'年报文件/{SINGLE_YEAR}/txt年报',
            target_year=SINGLE_YEAR,
            delete_pdf=DELETE_PDF,
            max_retries=MAX_RETRIES,
            timeout=TIMEOUT,
            processes=PROCESSES,
            upload_to_gdrive=UPLOAD_TO_GDRIVE,
            delete_local_after_upload=DELETE_LOCAL_AFTER_UPLOAD,
            gdrive_folder_id=GDRIVE_FOLDER_ID,
            gdrive_credentials_path=GDRIVE_CREDENTIALS_PATH,
            gdrive_token_path=GDRIVE_TOKEN_PATH,
            upload_record_file=UPLOAD_RECORD_FILE,
        )
        
        processor = AnnualReportProcessor(config)
        processor.run()
        
        print(f"\n{SINGLE_YEAR}年年报处理完毕\n")
