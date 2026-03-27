import configparser
from pathlib import Path
import requests
from fastapi import FastAPI, HTTPException, File, UploadFile, Body
from pydantic import BaseModel, Field
from typing import List, Optional, Dict, Any

# 从 conf/sc_web.conf 读取服务地址
_conf = configparser.ConfigParser()
_conf.read(Path(__file__).parent / "conf" / "sc_web.conf", encoding="utf-8")
TARGET_BASE_URL = _conf.get("docx_service", "base_url")
# 超时配置（秒）
TIMEOUT_CONFIG = {
    "split": 600,
    "merge": 600,
    "delete": 600
}


# -------------------------- 数据模型定义 --------------------------
class FileInfo(BaseModel):
    is_had_title: Optional[int] = None


class TreeItem(BaseModel):
    eid: str
    level: int
    idx: int
    children: Optional[List["TreeItem"]] = None
    text: Optional[str] = None
    id: Optional[int] = None
    parent_id: Optional[int] = None  # 父节点数据库ID，NULL表示根节点
    file_name: Optional[str] = None
    file_info: Optional[FileInfo] = None
    file_path: Optional[str] = None
    update_file_path: Optional[str] = ""
    node_type: Optional[str] = ""
    is_conversion_completion: Optional[int] = 0

    def __post_init__(self):
        # 强制兜底：无论传入什么，都确保children是列表
        if self.children is None or not isinstance(self.children, list):
            self.children = []



# 解决模型自引用问题
TreeItem.update_forward_refs()


class MergeRequest(BaseModel):
    tree: List[TreeItem]
    files: List[str]
    format_args: Dict[Any, Any]


class DeleteRequest(BaseModel):
    id: str = Field(..., description="使用方提供的唯一id")


class SplitResponse(BaseModel):
    status: int
    msg: str
    data: Dict[str, Any]

    def __init__(self, *args, **kwargs):
        # 先打印传入的参数
        print("===== 传入的参数信息 =====")
        print(f"位置参数 args: {args}")
        print(f"关键字参数 kwargs: {kwargs}")
        # 打印kwargs中每个变量的详细信息（可选，更清晰）
        if kwargs:
            print("kwargs 中的具体变量：")
            for key, value in kwargs.items():
                print(f"  {key} = {value} (类型: {type(value)})")

        # 必须将参数传给父类的__init__，否则BaseModel无法正常初始化
        super().__init__(*args, **kwargs)



class DeleteResponse(BaseModel):
    status: int
    msg: str
    data: Dict[str, Any]


# -------------------------- 核心接口调用函数 --------------------------
def call_docx_split(file_stream: bytes, file_name: str, file_id: str, had_title:int, rm_outline_in_doc:int) -> SplitResponse:
    """
    调用文件拆分接口（同步）
    :param file_stream: 文件字节流
    :param file_name: 原始文件名
    :param file_id: 唯一标识id
    :param had_title: 是否含有标题
    :param rm_outline_in_doc: 是否去掉html内部outline9
    :return: 拆分接口返回结果
    """
    url = f"{TARGET_BASE_URL}/api/tool_api/docx/split"
    try:
        # 构造multipart/form-data请求
        files = {
            "file": (file_name, file_stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        }
        data = {"id": file_id, "user_key": "DC4096F87722AD140F01AF8C3315B9A6", "had_title":had_title, "rm_outline_in_doc":rm_outline_in_doc}

        response = requests.post(
            url,
            files=files,
            data=data,
            timeout=TIMEOUT_CONFIG["split"]
        )
        response.raise_for_status()  # 抛出HTTP状态码异常
        result = response.json()
        print(result)
        return SplitResponse(**result)
    except requests.exceptions.HTTPError as e:
        raise HTTPException(
            status_code=e.response.status_code,
            detail=f"拆分接口调用失败: {e.response.text}"
        )
    except requests.exceptions.RequestException as e:
        raise HTTPException(
            status_code=500,
            detail=f"拆分接口网络异常: {str(e)}"
        )
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"拆分接口调用异常: {str(e)}"
        )


def call_docx_merge(merge_request: MergeRequest):
    """
    调用文件合并接口（同步）
    :param merge_request: 合并请求参数（tree+files）
    :return: 合并后的文件字节流
    """
    url = f"{TARGET_BASE_URL}/api/tool_api/docx/megre"  # 文档中拼写为 megre（merge笔误）
    try:
        data_ = merge_request.dict(exclude_unset=True)
        print(data_)
        data_["user_key"] = "DC4096F87722AD140F01AF8C3315B9A6"
        response = requests.post(
            url,
            json=data_,
            timeout=TIMEOUT_CONFIG["merge"]
        )
        response.raise_for_status()
        # 返回合并后的文件字节流
        return SplitResponse(**response.json())
    except requests.exceptions.HTTPError as e:
        raise HTTPException(
            status_code=e.response.status_code,
            detail=f"合并接口调用失败: {e.response.text}"
        )
    except requests.exceptions.RequestException as e:
        raise HTTPException(
            status_code=500,
            detail=f"合并接口网络异常: {str(e)}"
        )
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"合并接口调用异常: {str(e)}"
        )


def call_docx_delete(delete_request: DeleteRequest) -> DeleteResponse:
    """
    调用文件删除接口（同步）
    :param delete_request: 删除请求参数（id）
    :return: 删除接口返回结果
    """
    url = f"{TARGET_BASE_URL}/api/tool_api/docx/del"
    try:
        response = requests.post(
            url,
            json=delete_request.dict(),
            headers={"Content-Type": "application/json"},
            timeout=TIMEOUT_CONFIG["delete"]
        )
        response.raise_for_status()
        return DeleteResponse(**response.json())
    except requests.exceptions.HTTPError as e:
        raise HTTPException(
            status_code=e.response.status_code,
            detail=f"删除接口调用失败: {e.response.text}"
        )
    except requests.exceptions.RequestException as e:
        raise HTTPException(
            status_code=500,
            detail=f"删除接口网络异常: {str(e)}"
        )
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"删除接口调用异常: {str(e)}"
        )