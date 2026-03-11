import requests
from fastapi import FastAPI, HTTPException, File, UploadFile, Body
from pydantic import BaseModel, Field
from typing import List, Optional, Dict, Any

# 配置目标服务地址（根据实际情况修改）
TARGET_BASE_URL = "http://10.13.6.180:21001"
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
    file_name: Optional[str] = None
    file_info: Optional[FileInfo] = None
    file_path: Optional[str] = None
    update_file_path: Optional[str] = ""  # 更新后的文件路径
    node_type: Optional[str] = ""  # 节点类型：main/branch
    is_conversion_completion: Optional[int] = 0  # 是否转换完成

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
def call_docx_split(file_stream: bytes, file_name: str, file_id: str) -> SplitResponse:
    """
    调用文件拆分接口（同步）
    :param file_stream: 文件字节流
    :param file_name: 原始文件名
    :param file_id: 唯一标识id
    :return: 拆分接口返回结果
    """
    url = f"{TARGET_BASE_URL}/api/tool_api/docx/split"
    try:
        # 构造multipart/form-data请求
        files = {
            "file": (file_name, file_stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        }
        data = {"id": file_id, "user_key": "DC4096F87722AD140F01AF8C3315B9A6"}

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
    url = f"{TARGET_BASE_URL}/api/tool_api/docx/megre"  # 文档中为megre（merge笔误）
    try:
        data_ = merge_request.dict(exclude_unset=True)
        data_["user_key"] = "DC4096F87722AD140F01AF8C3315B9A6"
        response = requests.post(
            url,
            json=data_,
            timeout=TIMEOUT_CONFIG["split"]
        )
        # response = requests.post(
        #     url,
        #     json=merge_request.dict(exclude_unset=True),
        #     headers={"Content-Type": "application/json"},
        #     timeout=TIMEOUT_CONFIG["merge"]
        # )
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