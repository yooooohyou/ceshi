from spire.doc import *
from spire.doc.common import *
import os
import zipfile
import shutil
import re
import base64
import tempfile
import uuid  # 引入uuid生成唯一临时目录，避免冲突


class DocxHtmlConverter:
    """
    DOCX与HTML互转工具类（修复外部引用路径问题）
    核心改进：
    1. 所有路径强制使用绝对路径，脱离工作目录依赖
    2. 临时目录使用唯一ID命名，避免多线程/多调用冲突
    3. 图片路径匹配改为全量绝对路径匹配
    4. 增强路径校验和异常处理
    """

    def __init__(self):
        """初始化转换器，可自定义全局配置"""
        self.default_image_format = 0  # 0=PNG，1=JPG，2=BMP，3=GIF
        self.html_validation_type = XHTMLValidationType.none
        # 生成唯一临时目录前缀，避免冲突
        self.temp_dir_prefix = f"spire_temp_{uuid.uuid4().hex[:8]}"

    def _normalize_path(self, path):
        """【内部方法】统一路径格式并转为绝对路径"""
        if not path:
            return ""
        # 先转为绝对路径，再统一分隔符
        abs_path = os.path.abspath(path)
        return abs_path.replace('\\', '/').replace('//', '/')

    def _get_image_order_from_docx(self, docx_path):
        """【内部方法】解析DOCX，提取图片在文档中的显示顺序"""
        image_order = []
        try:
            # 1. 解析关系文件，建立ID→图片名映射
            rels_mapping = {}
            with zipfile.ZipFile(docx_path, 'r') as zip_file:
                if 'word/_rels/document.xml.rels' in zip_file.namelist():
                    rels_content = zip_file.read('word/_rels/document.xml.rels').decode('utf-8')
                    rel_pattern = re.compile(
                        r'<Relationship\s+Id="(rId\d+)"\s+Type="[^"]*image[^"]*"\s+Target="([^"]+)"')
                    for match in rel_pattern.finditer(rels_content):
                        r_id = match.group(1)
                        target = match.group(2)
                        img_name = os.path.basename(target)
                        rels_mapping[r_id] = img_name

            # 2. 解析文档内容，提取图片ID出现顺序
            with zipfile.ZipFile(docx_path, 'r') as zip_file:
                if 'word/document.xml' in zip_file.namelist():
                    xml_content = zip_file.read('word/document.xml').decode('utf-8')
                    id_pattern = re.compile(r'(embed|link|r:id|id)="(rId\d+)"')
                    blip_ids = []
                    for match in id_pattern.finditer(xml_content):
                        r_id = match.group(2)
                        if r_id not in blip_ids:
                            blip_ids.append(r_id)

            # 3. 转换ID为图片名
            for r_id in blip_ids:
                if r_id in rels_mapping and rels_mapping[r_id] not in image_order:
                    image_order.append(rels_mapping[r_id])

        except Exception as e:
            print(f"⚠️ 解析图片顺序失败：{e}，将使用文件名排序")
            with zipfile.ZipFile(docx_path, 'r') as zip_file:
                img_names = []
                for file_info in zip_file.infolist():
                    if file_info.filename.startswith('word/media/') and not file_info.is_dir():
                        img_names.append(os.path.basename(file_info.filename))
                image_order = sorted(img_names)

        print(f"✅ 解析到图片显示顺序：{image_order}")
        return image_order

    def _extract_original_images(self, docx_path, output_img_dir):
        """【内部方法】从DOCX中提取原始无压缩图片（强制绝对路径）"""
        # 确保输出目录是绝对路径
        output_img_dir = self._normalize_path(output_img_dir)

        # 清理并重建目录
        if os.path.exists(output_img_dir):
            shutil.rmtree(output_img_dir, ignore_errors=True)
        os.makedirs(output_img_dir, exist_ok=True)

        # 提取图片
        with zipfile.ZipFile(docx_path, 'r') as zip_file:
            for file_info in zip_file.infolist():
                if file_info.filename.startswith('word/media/') and not file_info.is_dir():
                    img_filename = os.path.basename(file_info.filename)
                    # 强制使用绝对路径保存
                    save_path = self._normalize_path(os.path.join(output_img_dir, img_filename))
                    with open(save_path, 'wb') as f:
                        f.write(zip_file.read(file_info.filename))

        print(f"✅ 提取原始图片完成，绝对路径：{output_img_dir}")
        return [f for f in os.listdir(output_img_dir) if os.path.isfile(os.path.join(output_img_dir, f))]

    def _image_to_base64(self, img_path):
        """【内部方法】将图片文件转为Base64编码（带MIME前缀）"""
        # 强制转为绝对路径
        img_path = self._normalize_path(img_path)

        try:
            if not os.path.exists(img_path):
                print(f"⚠️ 图片文件不存在（绝对路径）：{img_path}")
                return ""

            with open(img_path, 'rb') as f:
                img_data = f.read()

            # 根据后缀匹配MIME类型
            img_ext = os.path.splitext(img_path)[1].lower()
            mime_map = {
                '.png': 'image/png',
                '.jpg': 'image/jpeg',
                '.jpeg': 'image/jpeg',
                '.gif': 'image/gif',
                '.bmp': 'image/bmp'
            }
            mime_type = mime_map.get(img_ext, 'image/png')
            base64_str = f"data:{mime_type};base64,{base64.b64encode(img_data).decode('utf-8')}"
            return base64_str
        except Exception as e:
            print(f"⚠️ 图片 {img_path} 转Base64失败：{e}")
            return ""

    def docx_to_single_html(self, docx_path, html_path):
        """
        公开方法：DOCX转单文件HTML（修复外部引用路径问题）
        :param docx_path: 输入DOCX文件路径（支持相对/绝对）
        :param html_path: 输出HTML文件路径（支持相对/绝对）
        :return: 生成的HTML文本内容
        """
        # 1. 基础校验
        docx_path = self._normalize_path(docx_path)
        html_path = self._normalize_path(html_path)

        if not os.path.exists(docx_path):
            print(f"❌ 输入DOCX文件不存在（绝对路径）：{docx_path}")
            return ""

        # 确保输出目录存在
        html_dir = os.path.dirname(html_path)
        os.makedirs(html_dir, exist_ok=True)

        # 2. 创建唯一临时目录（基于HTML输出目录，绝对路径）
        spire_temp_dir = self._normalize_path(os.path.join(html_dir, self.temp_dir_prefix))
        original_img_dir = self._normalize_path(os.path.join(spire_temp_dir, "original_images"))
        spire_img_dir = self._normalize_path(os.path.join(spire_temp_dir, "images"))  # Spire生成图片的目录

        # 3. 获取图片顺序+提取原始图片
        image_display_order = self._get_image_order_from_docx(docx_path)
        extracted_imgs = self._extract_original_images(docx_path, original_img_dir)
        if not image_display_order and extracted_imgs:
            image_display_order = sorted(extracted_imgs)

        # 4. Spire转换生成临时HTML（带外部资源）
        document = Document()
        try:
            document.LoadFromFile(docx_path)
            # 配置导出选项（强制绝对路径）
            document.HtmlExportOptions.ImageEmbedded = False
            document.HtmlExportOptions.ImagesPath = spire_img_dir  # 绝对路径
            document.HtmlExportOptions.ImageFormat = self.default_image_format
            document.SaveToFile(html_path, FileFormat.Html)
        except Exception as e:
            print(f"❌ Spire转换HTML失败：{e}")
            return ""
        finally:
            document.Close()
            del document

        # 5. 读取并处理HTML内容
        with open(html_path, 'r', encoding='utf-8') as f:
            html_content = f.read()

        # 统一HTML中的路径分隔符
        html_content = html_content.replace('\\', '/')

        # 6. 提取Spire生成的图片列表（绝对路径匹配）
        spire_img_names = []
        img_pattern = re.compile(r'<img[^>]*src="([^"]+)"[^>]*>')
        for match in img_pattern.finditer(html_content):
            img_src = match.group(1)
            # 不管是相对还是绝对路径，只取文件名
            spire_img_name = os.path.basename(self._normalize_path(img_src))
            if spire_img_name not in spire_img_names:
                spire_img_names.append(spire_img_name)

        # 7. 内嵌CSS
        css_file_path = os.path.splitext(html_path)[0] + '_styles.css'
        css_file_path = self._normalize_path(css_file_path)
        if os.path.exists(css_file_path):
            with open(css_file_path, 'r', encoding='utf-8') as f:
                css_content = f.read()
            # 移除外部CSS引用，添加内联style
            html_content = re.sub(r'<link[^>]*href="[^"]+\.css"[^>]*>', '', html_content)
            html_content = html_content.replace(
                '</head>',
                f'<style type="text/css">\n{css_content}\n</style>\n</head>'
            )
            print(f"✅ 已内嵌CSS样式（CSS文件路径：{css_file_path}）")

        # 8. 内嵌图片（核心修复：绝对路径匹配+替换）
        for idx, spire_name in enumerate(spire_img_names):
            if idx < len(image_display_order):
                original_img_name = image_display_order[idx]
                # 强制使用绝对路径查找原始图片
                original_img_path = self._normalize_path(os.path.join(original_img_dir, original_img_name))
                base64_str = self._image_to_base64(original_img_path)

                if base64_str:
                    # 修复匹配规则：兼容相对/绝对路径的src属性
                    # 匹配所有包含该图片名的src属性（无论路径前缀）
                    spire_img_pattern = re.compile(f'src="[^"]*{re.escape(spire_name)}"')
                    html_content = spire_img_pattern.sub(f'src="{base64_str}"', html_content)
                    print(f"🔄 图片 {original_img_name} 已转为Base64内嵌（替换Spire生成的 {spire_name}）")
                else:
                    print(f"⚠️ 图片 {original_img_name} Base64转换失败，跳过替换")

        # 9. 保存最终HTML（强制覆盖）
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

        # 10. 清理临时文件（忽略删除失败的情况）
        temp_paths = [spire_temp_dir, css_file_path]
        for temp_path in temp_paths:
            if os.path.exists(temp_path):
                try:
                    if os.path.isdir(temp_path):
                        shutil.rmtree(temp_path, ignore_errors=True)
                    else:
                        os.remove(temp_path)
                    print(f"🗑️ 清理临时文件：{temp_path}")
                except Exception as e:
                    print(f"⚠️ 清理临时文件失败 {temp_path}：{e}")

        print(f"\n🎉 DOCX转HTML完成！")
        print(f"📄 最终文件绝对路径：{html_path}")
        print(f"✅ 特性：图片Base64内嵌 | CSS内嵌 | 图片无压缩 | 顺序对齐")
        return html_content

    def html_text_to_docx(self, html_text: str, output_docx_path: str):
        """
        公开方法：HTML文本转DOCX（兼容所有Spire.Doc版本，无临时文件残留）
        """
        # 强制转为绝对路径
        output_docx_path = self._normalize_path(output_docx_path)

        if not html_text.strip():
            print(f"❌ HTML文本为空，无法转换")
            return False

        document = None
        temp_html_file = None
        try:
            # 确保输出目录存在
            output_dir = os.path.dirname(output_docx_path)
            os.makedirs(output_dir, exist_ok=True)

            # 1. 创建唯一临时HTML文件（绝对路径）
            temp_html_file = tempfile.NamedTemporaryFile(
                mode='w',
                suffix='.html',
                delete=False,
                encoding='utf-8'
            )
            temp_html_path = self._normalize_path(temp_html_file.name)

            # 2. 写入HTML文本并关闭句柄
            temp_html_file.write(html_text)
            temp_html_file.close()

            # 3. 加载并转换
            document = Document()
            document.LoadFromFile(
                temp_html_path,
                FileFormat.Html,
                self.html_validation_type
            )
            document.SaveToFile(output_docx_path, FileFormat.Docx2016)

            print(f"✅ HTML转DOCX成功！文件绝对路径：{output_docx_path}")
            return True

        except Exception as e:
            print(f"❌ HTML转DOCX失败：{str(e)}")
            return False
        finally:
            # 强制清理资源
            if document:
                document.Close()
                del document
            if temp_html_file and os.path.exists(temp_html_file.name):
                try:
                    os.remove(temp_html_file.name)
                    # print(f"🗑️ 临时HTML文件已清理：{temp_html_file.name}")
                except Exception as e:
                    print(f"⚠️ 清理临时HTML文件失败：{e}")


# ------------------------------ 外部调用示例 ------------------------------
if __name__ == "__main__":
    # 初始化转换器
    converter = DocxHtmlConverter()

    # 示例1：DOCX转单文件HTML（支持外部调用）
    # 建议使用绝对路径，避免工作目录问题
    input_docx = r"C:\Users\you62\PyCharmMiscProject\word2html\input.docx"  # 替换为你的绝对路径
    output_html = r"C:\Users\you62\PyCharmMiscProject\word2html\output.html"  # 替换为你的绝对路径
    html_content = converter.docx_to_single_html(input_docx, output_html)

    # 示例2：HTML转DOCX
    if html_content:
        converter.html_text_to_docx(html_content, r"C:\Users\you62\PyCharmMiscProject\word2html\result.docx")
# # ------------------------------ 调用示例 ------------------------------
# if __name__ == "__main__":
#     # 1. 初始化转换器
#     converter = DocxHtmlConverter()
#
#     # 2. 示例1：DOCX转单文件HTML
#     input_docx = r"input.docx"  # 替换为你的DOCX路径
#     output_html = r"output.html"  # 输出HTML路径
#     html_content_a = converter.docx_to_single_html(input_docx, output_html)
#     # 写入最终HTML文件
#     with open(output_html, 'w', encoding='utf-8') as f:
#         f.write(html_content_a)
#
#     # 3. 示例2：DOCX转单文件HTML
#     sample_html = """自定义html"""
#
#     # 调用函数转换
#     converter.html_text_to_docx(html_content_a, "html_text_result.docx")