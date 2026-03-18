from spire.doc import *
from spire.doc.common import *
import os
import zipfile
import shutil
import re
import base64
import tempfile
import uuid


class DocxHtmlConverter:
    """
    DOCX与HTML互转工具类
    核心特性：
    1. 所有路径强制使用绝对路径，脱离工作目录依赖
    2. 临时目录使用唯一ID命名，避免多线程/多调用冲突
    3. 图片顺序解析覆盖正文、页眉、页脚、脚注、尾注等所有XML区域
    4. 降级兜底：Spire生成图片路径（find_actual_img_dir递归查找）
    5. 增强路径校验和异常处理
    """

    def __init__(self):
        """初始化转换器"""
        self.default_image_format = 0  # 0=PNG，1=JPG，2=BMP，3=GIF
        self.html_validation_type = XHTMLValidationType.none
        self.temp_dir_prefix = f"spire_temp_{uuid.uuid4().hex[:8]}"

    def _normalize_path(self, path):
        """【内部方法】统一路径格式并转为绝对路径"""
        if not path:
            return ""
        abs_path = os.path.abspath(path)
        return abs_path.replace('\\', '/').replace('//', '/')

    def _get_image_order_from_docx(self, docx_path):
        """
        【内部方法】解析DOCX，提取图片在文档中的显示顺序
        覆盖范围：正文、页眉、页脚、脚注、尾注等所有XML区域
        """
        image_order = []
        try:
            with zipfile.ZipFile(docx_path, 'r') as zip_file:
                all_files = zip_file.namelist()

                # 找出所有需要扫描的 XML 文件（正文 + 页眉 + 页脚 + 脚注 + 尾注）
                target_xml_files = [
                    f for f in all_files
                    if re.match(r'word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$', f)
                ]

                # 为每个 XML 文件加载对应的 rels，建立 rId → 图片名 的映射
                all_rels = {}
                for xml_file in target_xml_files:
                    xml_basename = os.path.basename(xml_file)
                    rels_path = f'word/_rels/{xml_basename}.rels'
                    if rels_path in all_files:
                        rels_content = zip_file.read(rels_path).decode('utf-8')
                        rel_pattern = re.compile(
                            r'<Relationship\s+Id="(rId\d+)"\s+Type="[^"]*image[^"]*"\s+Target="([^"]+)"'
                        )
                        rels_for_this_xml = {}
                        for match in rel_pattern.finditer(rels_content):
                            r_id, target = match.group(1), match.group(2)
                            rels_for_this_xml[r_id] = os.path.basename(target)
                        all_rels[xml_file] = rels_for_this_xml

                # 按 XML 文件顺序扫描图片引用（去重，保留首次出现顺序）
                seen = set()
                id_pattern = re.compile(r'(?:embed|link|r:id|id)="(rId\d+)"')
                for xml_file in target_xml_files:
                    if xml_file not in all_rels:
                        continue
                    xml_content = zip_file.read(xml_file).decode('utf-8')
                    rels_map = all_rels[xml_file]
                    for match in id_pattern.finditer(xml_content):
                        r_id = match.group(1)
                        if r_id in rels_map:
                            img_name = rels_map[r_id]
                            if img_name not in seen:
                                seen.add(img_name)
                                image_order.append(img_name)

        except Exception as e:
            print(f"⚠️ 解析图片顺序失败：{e}，将使用文件名排序")
            with zipfile.ZipFile(docx_path, 'r') as zip_file:
                image_order = sorted(
                    os.path.basename(f.filename)
                    for f in zip_file.infolist()
                    if f.filename.startswith('word/media/') and not f.is_dir()
                )

        print(f"✅ 解析到图片显示顺序：{image_order}")
        return image_order

    def _extract_original_images(self, docx_path, output_img_dir):
        """【内部方法】从DOCX中提取原始无压缩图片（强制绝对路径）"""
        output_img_dir = self._normalize_path(output_img_dir)

        if os.path.exists(output_img_dir):
            shutil.rmtree(output_img_dir, ignore_errors=True)
        os.makedirs(output_img_dir, exist_ok=True)

        with zipfile.ZipFile(docx_path, 'r') as zip_file:
            for file_info in zip_file.infolist():
                if file_info.filename.startswith('word/media/') and not file_info.is_dir():
                    img_filename = os.path.basename(file_info.filename)
                    save_path = self._normalize_path(os.path.join(output_img_dir, img_filename))
                    with open(save_path, 'wb') as f:
                        f.write(zip_file.read(file_info.filename))

        return [f for f in os.listdir(output_img_dir) if os.path.isfile(os.path.join(output_img_dir, f))]

    def _find_actual_img_dir(self, base_dir):
        """
        【内部方法】递归查找第一个实际包含图片文件的目录
        用于兜底：Spire有时将图片输出到嵌套子目录中
        """
        img_exts = ('.png', '.jpg', '.jpeg', '.gif', '.bmp')
        for root, dirs, files in os.walk(base_dir):
            if any(f.lower().endswith(img_exts) for f in files):
                return root
        return base_dir

    def _image_to_base64(self, img_path):
        """【内部方法】将图片文件转为Base64编码（带MIME前缀）"""
        img_path = self._normalize_path(img_path)
        try:
            if not os.path.exists(img_path):
                print(f"⚠️ 图片文件不存在（绝对路径）：{img_path}")
                return ""

            with open(img_path, 'rb') as f:
                img_data = f.read()

            img_ext = os.path.splitext(img_path)[1].lower()
            mime_map = {
                '.png': 'image/png',
                '.jpg': 'image/jpeg',
                '.jpeg': 'image/jpeg',
                '.gif': 'image/gif',
                '.bmp': 'image/bmp'
            }
            mime_type = mime_map.get(img_ext, 'image/png')
            return f"data:{mime_type};base64,{base64.b64encode(img_data).decode('utf-8')}"

        except Exception as e:
            print(f"⚠️ 图片 {img_path} 转Base64失败：{e}")
            return ""

    def docx_to_single_html(self, docx_path, html_path):
        """
        公开方法：DOCX转单文件HTML
        特性：图片Base64内嵌 | CSS内嵌 | 图片无压缩 | 顺序对齐（覆盖所有XML区域）
        :param docx_path: 输入DOCX文件路径（支持相对/绝对）
        :param html_path: 输出HTML文件路径（支持相对/绝对）
        :return: 生成的HTML文本内容，失败返回空字符串
        """
        # 1. 路径校验与标准化
        docx_path = self._normalize_path(docx_path)
        html_path = self._normalize_path(html_path)

        if not os.path.exists(docx_path):
            print(f"❌ 输入DOCX文件不存在（绝对路径）：{docx_path}")
            return ""

        html_dir = os.path.dirname(html_path)
        os.makedirs(html_dir, exist_ok=True)

        # 2. 创建唯一临时目录
        spire_temp_dir = self._normalize_path(os.path.join(html_dir, self.temp_dir_prefix))
        original_img_dir = self._normalize_path(os.path.join(spire_temp_dir, "original_images"))
        spire_img_dir = self._normalize_path(os.path.join(spire_temp_dir, "images"))

        # 3. 解析图片顺序 + 提取原始图片
        image_display_order = self._get_image_order_from_docx(docx_path)
        extracted_imgs = self._extract_original_images(docx_path, original_img_dir)

        # 顺序解析为空时兜底
        if not image_display_order and extracted_imgs:
            image_display_order = sorted(extracted_imgs)
            print(f"⚠️ 顺序解析为空，兜底使用：{image_display_order}")

        # 4. Spire转换生成临时HTML
        document = Document()
        try:
            document.LoadFromFile(docx_path)
            document.HtmlExportOptions.ImageEmbedded = False
            document.HtmlExportOptions.ImagesPath = spire_img_dir
            document.HtmlExportOptions.ImageFormat = self.default_image_format
            document.SaveToFile(html_path, FileFormat.Html)
        except Exception as e:
            print(f"❌ Spire转换HTML失败：{e}")
            return ""
        finally:
            document.Close()
            del document

        # 5. 读取HTML，统一路径分隔符
        with open(html_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
        html_content = html_content.replace('\\', '/')

        # 6. 内嵌CSS
        css_file_path = self._normalize_path(os.path.splitext(html_path)[0] + '_styles.css')
        if os.path.exists(css_file_path):
            with open(css_file_path, 'r', encoding='utf-8') as f:
                css_content = f.read()
            html_content = re.sub(r'<link[^>]*href="[^"]+\.css"[^>]*>', '', html_content)
            html_content = html_content.replace(
                '</head>',
                f'<style type="text/css">\n{css_content}\n</style>\n</head>'
            )
            print("✅ 已内嵌CSS样式")

        # 7. 提取Spire生成的图片文件名列表
        img_pattern = re.compile(r'<img[^>]*src="([^"]+)"[^>]*>')
        spire_img_names = []
        for match in img_pattern.finditer(html_content):
            spire_img_name = os.path.basename(self._normalize_path(match.group(1)))
            if spire_img_name not in spire_img_names:
                spire_img_names.append(spire_img_name)
        print(f"=== Spire 图片列表：{spire_img_names} ===")

        # 8. 内嵌图片（优先原始图片，数量不足时降级到Spire生成图片）
        actual_spire_img_dir = self._find_actual_img_dir(spire_img_dir)
        for idx, spire_name in enumerate(spire_img_names):
            if idx < len(image_display_order):
                # 优先：原始无压缩图片
                img_path = self._normalize_path(
                    os.path.join(original_img_dir, image_display_order[idx])
                )
            else:
                # 降级：Spire生成的图片（递归查找实际目录）
                img_path = self._normalize_path(
                    os.path.join(actual_spire_img_dir, spire_name)
                )

            if not os.path.exists(img_path):
                print(f"⚠️ 找不到图片：{img_path}，跳过")
                continue

            base64_str = self._image_to_base64(img_path)
            if not base64_str:
                continue

            html_content = re.compile(
                f'src="[^"]*{re.escape(spire_name)}"'
            ).sub(f'src="{base64_str}"', html_content)
            print(f"🔄 {spire_name} → {os.path.basename(img_path)} 已转为Base64")

        # 9. 保存最终HTML
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

        # 10. 清理临时文件
        for temp_path in [spire_temp_dir, css_file_path]:
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
        :param html_text: 输入HTML字符串
        :param output_docx_path: 输出DOCX文件路径（支持相对/绝对）
        :return: 成功返回True，失败返回False
        """
        output_docx_path = self._normalize_path(output_docx_path)

        if not html_text.strip():
            print("❌ HTML文本为空，无法转换")
            return False

        document = None
        temp_html_file = None
        try:
            output_dir = os.path.dirname(output_docx_path)
            os.makedirs(output_dir, exist_ok=True)

            # 创建唯一临时HTML文件
            temp_html_file = tempfile.NamedTemporaryFile(
                mode='w',
                suffix='.html',
                delete=False,
                encoding='utf-8'
            )
            temp_html_path = self._normalize_path(temp_html_file.name)
            temp_html_file.write(html_text)
            temp_html_file.close()

            document = Document()
            document.LoadFromFile(
                temp_html_path,
                FileFormat.Html,
                self.html_validation_type
            )
            document.SaveToFile(output_docx_path, FileFormat.Docx2016)
            return True

        except Exception as e:
            print(f"❌ HTML转DOCX失败：{str(e)}")
            return False
        finally:
            if document:
                document.Close()
                del document
            if temp_html_file and os.path.exists(temp_html_file.name):
                try:
                    os.remove(temp_html_file.name)
                except Exception as e:
                    print(f"⚠️ 清理临时HTML文件失败：{e}")


# ------------------------------ 调用示例 ------------------------------
if __name__ == "__main__":
    converter = DocxHtmlConverter()

    # 示例1：DOCX转单文件HTML
    # input_docx = r"inputzk1.docx"
    # output_html = r"output.html"
    # html_content = converter.docx_to_single_html(input_docx, output_html)

    # 示例2：HTML文本转DOCX
    # sample_html = """<html><body><p>Hello World</p></body></html>"""
    # converter.html_text_to_docx(sample_html, "output.docx")
    pass