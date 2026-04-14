import logging
import os
from typing import Any, List, Optional

from fastapi import APIRouter, Body, HTTPException

from app.models.schemas import unified_response
from app.utils.image_utils import download_image, download_images, generate_and_convert_to_html
from app.utils.path_utils import save_html_and_get_url
from docxautogenerator import generate_car_info_doc, generate_fully_centered_patent_doc, generate_report_doc

router = APIRouter(prefix="/generate_patent_doc")
logger = logging.getLogger(__name__)


@router.post("/default", summary="生成器示例接口")
async def generate_default_patent_doc():
    """使用默认的专利数据和图片路径生成文档并返回下载"""
    from app.core.config import UPLOAD_DIR
    from app.converters.docx_converter import docx_to_html
    try:
        save_path = os.path.join(UPLOAD_DIR, "专利报告.docx")
        DEFAULT_PATENT_DATA = [
            ["1", "发明专利", "一种核岛用防爆电话电缆", "ZL201210300077.1", "远程电缆股份有限公司", "2015-07-29"],
            ["2", "发明专利", "一种电缆保护夹座安装方法", "ZL201110454443.4", "远程电缆股份有限公司", "2014-06-25"],
            ["3", "发明专利", "硅烷交联聚乙烯绝缘电缆料及其制造方法", "ZL200710022494.3", "远程电缆股份有限公司", "2010-12-08"],
        ]
        DEFAULT_CERT_IMG_PATHS = [
            "./pic/图片2.jpg", "./pic/图片3.jpg", "./pic/图片4.jpg",
            "./pic/图片5.jpg", "./pic/图片6.jpg", "./pic/图片7.jpg",
        ]
        generate_fully_centered_patent_doc(
            DEFAULT_PATENT_DATA, DEFAULT_CERT_IMG_PATHS, save_path, last_img_display_mode=1
        )
        if not os.path.exists(save_path):
            raise HTTPException(status_code=404, detail="文档生成失败，文件不存在")
        html_content, _ = docx_to_html(save_path)
        try:
            if os.path.exists(save_path):
                os.remove(save_path)
        except Exception as e:
            logger.warning(f"警告：无法删除临时文件 {save_path} - {e}")
        return unified_response(200, "节点HTML内容更新成功", {"http_path": save_html_and_get_url(html_content)})
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"生成文档时出错: {str(e)}")


@router.post(
    "/patent_generator",
    summary="专利生成器",
    description="生成包含专利表格和证书图片的Word文档，并转换为单文件HTML",
)
async def patent_generator(
    patent_data: List[List[Any]] = Body(
        ...,
        description="专利数据二维列表，每行：[序号,专利类型,专利名称,专利号,专利权人,授权公告日]",
        min_items=1,
        max_items=100,
    ),
    cert_img_urls: List[Any] = Body([], description="专利证书图片URL列表（支持多层嵌套）"),
    fill_empty_space: bool = Body(False, description="是否自动补图填充页面空白区域"),
    last_img_display_mode: int = Body(1, description="奇数图片时最后一行的显示模式：1=删除空单元格，2=合并单元格", ge=1, le=2),
):
    def _flatten_urls(obj) -> List[str]:
        if obj is None:
            return []
        if isinstance(obj, str):
            return [obj] if obj.strip() else []
        if isinstance(obj, list):
            result = []
            for item in obj:
                result.extend(_flatten_urls(item))
            return result
        return [str(obj)] if obj else []

    try:
        for idx, row in enumerate(patent_data):
            if len(row) != 6:
                raise ValueError(
                    f"专利数据第{idx + 1}行必须包含6个字段（序号、专利类型、专利名称、专利号、专利权人、授权公告日）"
                )

        flat_urls = _flatten_urls(cert_img_urls)
        cert_img_paths = await download_images(flat_urls)

        html_content = generate_and_convert_to_html(
            generate_fully_centered_patent_doc,
            patent_data=patent_data,
            cert_img_paths=cert_img_paths,
            fill_empty_space=fill_empty_space,
            last_img_display_mode=last_img_display_mode,
        )

        try:
            for img_path in cert_img_paths:
                if img_path:
                    os.remove(img_path)
        except Exception:
            pass

        if not html_content:
            raise HTTPException(status_code=500, detail="HTML转换失败，生成的HTML内容为空")

        return unified_response(200, "生成专利生成器HTML内容成功", {"http_path": save_html_and_get_url(html_content)})

    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"生成专利文档失败：{str(e)}")


@router.post(
    "/financial_report_generator",
    summary="财报生成器",
    description="生成包含标题和多张图片的年度报告Word文档，并转换为单文件HTML返回",
)
async def financial_report_generator(
    title_text: str = Body(..., description="报告标题文本", min_length=1, max_length=100),
    second_row_img_url: str = Body(..., description="第二行主图片的网络URL（必填）"),
    other_img_urls: List[str] = Body([], description="其他图片的网络URL列表（可选）"),
    columns_per_row: int = Body(2, description="每行显示的图片列数", ge=1, le=4),
):
    try:
        main_img_path = await download_image(second_row_img_url)
        other_img_paths = await download_images(other_img_urls)

        html_content = generate_and_convert_to_html(
            generate_report_doc,
            title_text=title_text,
            second_row_img_path=main_img_path,
            other_img_paths=other_img_paths,
            columns_per_row=columns_per_row,
        )

        try:
            if main_img_path:
                os.remove(main_img_path)
            for img_path in other_img_paths:
                if img_path:
                    os.remove(img_path)
        except Exception:
            pass

        if not html_content:
            raise HTTPException(status_code=500, detail="HTML转换失败，生成的HTML内容为空")

        return unified_response(200, "生成财报生成器HTML内容成功", {"http_path": save_html_and_get_url(html_content)})

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"生成报告失败：{str(e)}")


@router.post(
    "/vehicle_generator",
    summary="车辆生成器",
    description="生成包含公司车辆信息的Word文档，并转换为单文件HTML",
)
async def vehicle_generator(
    car_data: List[List[Any]] = Body(
        ...,
        description="车辆数据列表，每行：[序号,车牌号,行驶证URL,车辆图片URL]",
        min_items=1,
    ),
    table_title: str = Body("公司车辆信息", description="表格标题", min_length=1, max_length=50),
):
    try:
        all_img_urls = []
        for car_row in car_data:
            seq, plate_num, drive_url, car_url = car_row
            all_img_urls.extend([drive_url, car_url])

        img_paths = await download_images(all_img_urls)
        img_path_map = dict(zip(all_img_urls, img_paths))

        car_data_with_local_paths = []
        for car_row in car_data:
            seq, plate_num, drive_url, car_url = car_row
            car_data_with_local_paths.append([
                seq, plate_num,
                img_path_map.get(drive_url, ""),
                img_path_map.get(car_url, ""),
            ])

        html_content = generate_and_convert_to_html(
            generate_car_info_doc,
            car_data=car_data_with_local_paths,
            table_title=table_title,
        )

        try:
            for img_path in img_paths:
                if img_path:
                    os.remove(img_path)
        except Exception:
            pass

        if not html_content:
            raise HTTPException(status_code=500, detail="HTML转换失败")

        return unified_response(200, "生成车辆生成器HTML内容成功", {"http_path": save_html_and_get_url(html_content)})

    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
