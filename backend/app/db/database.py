import datetime
import logging
import os
from contextlib import contextmanager
from typing import Any, Dict, List, Optional

import psycopg2
from psycopg2.extras import RealDictCursor

from app.core.config import POSTGRES_CONFIG
from app.core.constants import DEFAULT_MAIN_NODE

logger = logging.getLogger(__name__)


# ─── 连接管理 ────────────────────────────────────────────────────────────────

@contextmanager
def get_db_connection():
    """PostgreSQL 数据库连接上下文管理器"""
    conn = None
    try:
        conn = psycopg2.connect(**POSTGRES_CONFIG)
        conn.autocommit = False
        yield conn
    except Exception as e:
        if conn:
            conn.rollback()
        raise Exception(f"数据库操作异常：{str(e)}")
    finally:
        if conn:
            conn.close()


def init_db_tables():
    """初始化 PostgreSQL 数据表（首次运行时调用）"""
    create_file_table_sql = """
    CREATE TABLE IF NOT EXISTS "yxdl_docx_upload_records" (
      "id" SERIAL PRIMARY KEY,
      "original_filename" varchar(255) NOT NULL,
      "new_filename" varchar(255) NOT NULL,
      "save_path" varchar(512) NOT NULL,
      "upload_time" TIMESTAMP NOT NULL,
      "update_time" TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
      "split_file_id" varchar(128),
      "process_mode" varchar(16) DEFAULT 'split'
    );
    """

    create_title_tree_table_sql = """
    CREATE TABLE IF NOT EXISTS "yxdl_docx_title_trees" (
      "id" SERIAL PRIMARY KEY,
      "record_id" int4,
      "title_text" varchar(512) NOT NULL,
      "html_content" text,
      "create_time" TIMESTAMP NOT NULL,
      "update_time" TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
      "level" int4,
      "eid" varchar(128),
      "idx" int4,
      "node_type" varchar(16) DEFAULT 'main',
      "split_id" int4
    );
    """

    # 嵌入组件表：增量创建，不覆盖已有数据
    create_embed_components_sql = """
    CREATE TABLE IF NOT EXISTS "yxdl_embed_components" (
      "id"          SERIAL PRIMARY KEY,
      "embed_id"    varchar(64) UNIQUE NOT NULL,
      "embed_type"  varchar(32) NOT NULL,
      "version"     int4 NOT NULL DEFAULT 1,
      "title"       varchar(512) DEFAULT '',
      "display"     varchar(16) DEFAULT 'inline',
      "url"         varchar(1024),
      "payload"     jsonb NOT NULL DEFAULT '{}'::jsonb,
      "record_id"   int4,
      "node_id"     int4,
      "status"      int2 DEFAULT 1,
      "create_time" TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
      "update_time" TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
    );
    CREATE INDEX IF NOT EXISTS idx_embed_record ON "yxdl_embed_components" ("record_id");
    CREATE INDEX IF NOT EXISTS idx_embed_node   ON "yxdl_embed_components" ("node_id");
    CREATE INDEX IF NOT EXISTS idx_embed_type   ON "yxdl_embed_components" ("embed_type");
    """

    create_xlsx_upload_records_sql = """
    CREATE TABLE IF NOT EXISTS "yxdl_xlsx_upload_records" (
      "id"                SERIAL PRIMARY KEY,
      "original_filename" varchar(255) NOT NULL,
      "new_filename"      varchar(255) NOT NULL,
      "file_sign"         varchar(128) NOT NULL,
      "save_path"         varchar(512) NOT NULL,
      "upload_time"       TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
      "update_time"       TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
    );
    CREATE INDEX IF NOT EXISTS idx_xlsx_new_filename ON "yxdl_xlsx_upload_records" ("new_filename");
    CREATE INDEX IF NOT EXISTS idx_xlsx_file_sign    ON "yxdl_xlsx_upload_records" ("file_sign");
    """

    alter_upload_records_sql = """
    ALTER TABLE "yxdl_docx_upload_records"
    ADD COLUMN "title_font_dict" jsonb DEFAULT NULL;
    """

    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(create_file_table_sql)
                cursor.execute(create_title_tree_table_sql)
                cursor.execute(create_embed_components_sql)
                cursor.execute(create_xlsx_upload_records_sql)
                cursor.execute(alter_upload_records_sql)
                conn.commit()
        logger.debug("PostgreSQL数据表初始化成功")
    except Exception as e:
        logger.warning(f"警告：数据库初始化失败，部分功能将不可用：{e}")


# ─── 辅助函数 ────────────────────────────────────────────────────────────────

def deduplicate_dict_list(dict_list: list) -> list:
    """对包含字典的列表按 key-value 去重，保留首次出现"""
    seen = set()
    result = []
    for d in dict_list:
        key = tuple(sorted(d.items()))
        if key not in seen:
            seen.add(key)
            result.append(d)
    return result


def build_eid_path_mapping(files: List[str]) -> Dict[str, str]:
    """构建 eid → 文件路径 的映射（文件名不含后缀作为 eid）"""
    return {
        os.path.splitext(os.path.basename(fp))[0]: fp
        for fp in files
    }


def assign_file_path_to_tree(node, eid_path_map: Dict[str, str]):
    """递归为树节点分配对应的文件路径"""
    if node.eid in eid_path_map:
        node.file_path = eid_path_map[node.eid]
    else:
        logger.warning(f"assign_file_path_to_tree: eid={node.eid!r} 未在 files 中找到匹配")
    for child in (node.children or []):
        assign_file_path_to_tree(child, eid_path_map)


def build_simplified_tree(rows) -> list:
    """将数据库行列表按 parent_id 组织成嵌套树结构"""
    items = [dict(row) for row in rows]
    for item in items:
        item["children"] = []

    id_map = {item["id"]: item for item in items}
    tree = []
    for item in sorted(items, key=lambda x: x["idx"]):
        pid = item.get("parent_id")
        if pid and pid in id_map:
            id_map[pid]["children"].append(item)
        else:
            tree.append(item)

    def _sort_key(x):
        return (-(x.get("batch_count") or 0), x["idx"])

    def sort_children(node):
        node["children"].sort(key=_sort_key)
        for child in node["children"]:
            sort_children(child)

    fake_root = {"children": tree}
    sort_children(fake_root)
    tree.sort(key=_sort_key)
    return tree


def get_next_batch_count(record_id: int) -> int:
    """查询 record_id 下已有的最大 batch_count，返回 +1 后的值（首次返回 1）"""
    sql = """
        SELECT COALESCE(MAX(batch_count), 0) + 1
        FROM "yxdl_docx_title_trees"
        WHERE record_id = %s;
    """
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(sql, (record_id,))
            return cursor.fetchone()[0]


def get_tree_node_file_paths(record_id: int) -> List[str]:
    """查询 record_id 下所有节点的文件路径（优先 update_file_path）"""
    if not isinstance(record_id, int) or record_id <= 0:
        raise ValueError("record_id必须是正整数")

    select_sql = """
    SELECT
        CASE
            WHEN is_conversion_completion = 1 AND update_file_path IS NOT NULL AND update_file_path != ''
            THEN update_file_path
            ELSE origin_file_path
        END AS file_path
    FROM "yxdl_docx_title_trees"
    WHERE record_id = %s;
    """
    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(select_sql, (record_id,))
                raw_paths = [row[0] for row in cursor.fetchall()]
    except Exception as e:
        raise RuntimeError(f"查询文件路径失败：{str(e)}") from e

    seen = set()
    unique_paths = []
    for path in raw_paths:
        if path and isinstance(path, str) and path.strip() and path not in seen:
            seen.add(path)
            unique_paths.append(path)
    return unique_paths


# ─── 树节点处理 ──────────────────────────────────────────────────────────────

def process_split_tree_nodes(
    nodes,
    record_id: int,
    current_time: datetime.datetime,
    file_base_path: str,
    convert_html: bool = True,
    parent_id: Optional[int] = None,
    batch_count: int = 1,
) -> List[Dict[str, Any]]:
    """递归处理拆分后的树节点，入库并返回带层级结构的节点信息"""
    from mergfile import TreeItem
    from app.converters.docx_converter import docx_to_html

    if not isinstance(nodes, list) or not nodes:
        return []
    if not isinstance(record_id, int) or record_id <= 0:
        raise ValueError("record_id必须是正整数")

    result_nodes = []
    for node in nodes:
        if not isinstance(node, TreeItem):
            logger.warning(f"process_split_tree_nodes: 跳过非TreeItem节点 type={type(node)}")
            continue

        node_title = (
            node.text.strip()
            if node.text and isinstance(node.text, str)
            else f"节点_{node.eid or '未知'}_{node.level}_{node.idx}"
        )
        node_file_path = node.file_path or ""
        level = node.level

        if convert_html and node_file_path:
            try:
                html_content, _ = docx_to_html(node_file_path)
            except Exception as e:
                logger.error(f"process_split_tree_nodes: docx_to_html 失败 path={node_file_path} err={e}")
                html_content = ""
        else:
            if convert_html and not node_file_path:
                logger.warning(f"process_split_tree_nodes: 节点 eid={node.eid} file_path 为空，跳过转换")
            html_content = ""
        is_conversion_completion = 1 if html_content else 0

        insert_tree_sql = """
        INSERT INTO "yxdl_docx_title_trees"
        (record_id, title_text, html_content, create_time, update_time,
         level, eid, idx, node_type, origin_file_path, is_conversion_completion, parent_id, batch_count, split_id)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        RETURNING id;
        """
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(insert_tree_sql, (
                    record_id, node_title, html_content,
                    current_time, current_time,
                    node.level, node.eid, node.idx,
                    "branch", node_file_path, is_conversion_completion,
                    parent_id, batch_count, node.id
                ))
                node_id = cursor.fetchone()[0]
                conn.commit()
        logger.info(f"process_split_tree_nodes: 插入节点 id={node_id} eid={node.eid} level={level}")

        current_node = {
            "name": node_title,
            "node_id": node_id,
            "level": level,
            "file_name": node_file_path,
            "children": [],
        }
        children = node.children or []
        if children:
            current_node["children"] = process_split_tree_nodes(
                children, record_id, current_time, file_base_path,
                convert_html=convert_html,
                parent_id=node_id,
                batch_count=batch_count,
            )
        result_nodes.append(current_node)
    return result_nodes


def process_split_tree_nodes_with_select(
    tree_nodes_org,
    record_id: int,
    current_time: datetime.datetime,
    file_base_path: str,
) -> List[Dict[str, Any]]:
    """递归处理从数据库查询的树节点（无数据库插入操作）"""
    from mergfile import TreeItem

    if not isinstance(tree_nodes_org, list) or not tree_nodes_org:
        return []
    if not isinstance(record_id, int) or record_id <= 0:
        raise ValueError("record_id必须是正整数")

    result_nodes = []
    for node in tree_nodes_org:
        if not isinstance(node, TreeItem):
            continue
        try:
            node_title = (
                node.text.strip()
                if node.text and isinstance(node.text, str)
                else f"节点_{node.eid or '未知'}_{node.level}_{node.idx}"
            )
            node_file_path = node.file_path or ""
            level = node.level

            if node.id is not None:
                node_id = node.id
            else:
                unique_identifier = f"{record_id}_{node.level}_{node.idx}_{node.eid or 'no_eid'}"
                node_id = hash(unique_identifier) % (10 ** 8)

            current_node = {
                "name": node_title,
                "node_id": node_id,
                "level": level,
                "file_name": node_file_path,
                "update_file_path": node.update_file_path,
                "is_conversion_completion": node.is_conversion_completion,
                "children": [],
            }
            if node.children:
                current_node["children"] = process_split_tree_nodes_with_select(
                    node.children, record_id, current_time, file_base_path
                )
            result_nodes.append(current_node)
        except Exception:
            continue
    return result_nodes


def process_single_tree_node(
    node,
    record_id: int,
    id_: int,
    current_time: datetime.datetime,
    convert_html: bool = True,
) -> Dict[str, Any]:
    """处理单个树节点，更新数据库，返回节点信息"""
    from mergfile import TreeItem
    from app.converters.docx_converter import docx_to_html

    result_node = {
        "name": "",
        "node_id": None,
        "level": node.level,
        "file_name": node.file_path,
        "update_success": False,
    }

    if not isinstance(node, TreeItem):
        result_node["name"] = "无效节点类型"
        return result_node
    if not isinstance(record_id, int) or record_id <= 0:
        raise ValueError("record_id必须是正整数")

    try:
        node_title = (
            node.text.strip()
            if node.text and isinstance(node.text, str)
            else f"节点_{node.eid or '未知'}_{node.level}_{node.idx}"
        )

        if convert_html and node.file_path:
            try:
                node_html_content, _ = docx_to_html(node.file_path)
            except Exception:
                node_html_content = ""
        else:
            node_html_content = ""
        is_conv_done = 1 if node_html_content else 0

        if convert_html:
            update_tree_sql = """
            UPDATE "yxdl_docx_title_trees"
            SET update_time = %s, level = %s, origin_file_path = %s,
                html_content = %s, is_conversion_completion = %s, eid = %s
            WHERE record_id = %s AND id = %s
            RETURNING id;
            """
            params = (
                current_time, node.level, node.file_path,
                node_html_content, is_conv_done, node.eid,
                record_id, id_,
            )
        else:
            update_tree_sql = """
            UPDATE "yxdl_docx_title_trees"
            SET update_time = %s, level = %s, origin_file_path = %s,
                is_conversion_completion = %s, eid = %s
            WHERE record_id = %s AND id = %s
            RETURNING id;
            """
            params = (
                current_time, node.level, node.file_path,
                0, node.eid,
                record_id, id_,
            )

        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(update_tree_sql, params)
                update_result = cursor.fetchone()
                conn.commit()
                if update_result:
                    result_node["node_id"] = update_result[0]
                    result_node["update_success"] = True
                else:
                    logger.warning(f"process_single_tree_node: 未匹配到节点 record_id={record_id} id={id_}")

        result_node["name"] = node_title
    except ValueError:
        pass
    except Exception as e:
        logger.error(f"process_single_tree_node: 失败 eid={node.eid} err={e}")
    return result_node


def create_single_main_node(
    record_id: int,
    current_time: datetime.datetime,
    file_path: str,
) -> int:
    """创建单个主节点，返回节点 ID"""
    from app.converters.docx_converter import docx_to_html

    html_content, _ = docx_to_html(file_path)

    insert_tree_sql = """
    INSERT INTO "yxdl_docx_title_trees"
    (record_id, title_text, html_content, create_time, update_time,
     level, eid, idx, node_type, parent_id, batch_count)
    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    RETURNING id;
    """
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(insert_tree_sql, (
                record_id,
                DEFAULT_MAIN_NODE["title"],
                html_content,
                current_time,
                current_time,
                DEFAULT_MAIN_NODE["level"],
                DEFAULT_MAIN_NODE["eid"],
                DEFAULT_MAIN_NODE["idx"],
                "main",
                None,
                1,
            ))
            node_id = cursor.fetchone()[0]
            conn.commit()
    logger.debug(f"成功创建主节点：ID={node_id}, 标题={DEFAULT_MAIN_NODE['title']}")
    return node_id


def recover_split_tree_nodes(record_id: int) -> List[Dict[str, Any]]:
    """从数据库恢复拆分树节点，返回嵌套结构"""
    from mergfile import TreeItem

    if not isinstance(record_id, int) or record_id <= 0:
        raise ValueError("record_id必须是正整数")

    select_sql = """
    SELECT
        id, title_text, level, eid, idx, parent_id, batch_count,
        origin_file_path, update_file_path, is_conversion_completion, split_id
    FROM "yxdl_docx_title_trees"
    WHERE record_id = %s
    ORDER BY level ASC, idx ASC;
    """
    try:
        with get_db_connection() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cursor:
                cursor.execute(select_sql, (record_id,))
                node_records = cursor.fetchall()
    except Exception as e:
        raise RuntimeError(f"查询数据库失败：{str(e)}") from e

    if not node_records:
        return []

    nested = build_simplified_tree(node_records)

    def _remap(nodes: List[Dict]) -> List[Dict]:
        result = []
        for node in nodes:
            if node.get("is_conversion_completion") == 1 and node.get("update_file_path"):
                file_path = node["update_file_path"]
            else:
                file_path = node.get("origin_file_path") or ""
            file_name = os.path.splitext(os.path.basename(file_path))[0] if file_path else ""
            result.append({
                "eid":                      node.get("eid") or "",
                "text":                     node.get("title_text") or "",
                "level":                    node.get("level", 0),
                "id":                       node.get("split_id"),
                "idx":                      node.get("idx", 0),
                "parent_id":                node.get("parent_id"),
                "file_name":                file_name,
                "file_path":                file_path,
                "file_info":                {"is_had_title": 1},
                "update_file_path":         node.get("update_file_path") or "",
                "node_type":                node.get("node_type") or "",
                "is_conversion_completion": node.get("is_conversion_completion", 0),
                "children":                 _remap(node.get("children", [])),
            })
        return result

    return _remap(nested)


def query_and_build_tree(rec_id: int, cur_time: datetime.datetime) -> List[Dict[str, Any]]:
    """查询 record_id 下所有节点，组装成嵌套树结构后返回标准格式（供路由直接调用）"""
    from mergfile import TreeItem

    select_sql = """
        SELECT
            id, title_text, level, eid, idx, parent_id, batch_count,
            origin_file_path, update_file_path, is_conversion_completion
        FROM "yxdl_docx_title_trees"
        WHERE record_id = %s
        ORDER BY level ASC, idx ASC;
    """
    try:
        with get_db_connection() as conn:
            with conn.cursor(cursor_factory=RealDictCursor) as cursor:
                cursor.execute(select_sql, (rec_id,))
                node_records = cursor.fetchall()
    except Exception as e:
        raise RuntimeError(f"查询数据库失败：{str(e)}") from e

    tree_nodes_org = [
        TreeItem(**{
            "id":                       item.get("id"),
            "text":                     item.get("title_text"),
            "level":                    item.get("level"),
            "eid":                      item.get("eid"),
            "idx":                      item.get("idx"),
            "parent_id":                item.get("parent_id"),
            "file_path":                item.get("origin_file_path"),
            "update_file_path":         item.get("update_file_path", ""),
            "is_conversion_completion": item.get("is_conversion_completion", 0),
            "children":                 [],
            "file_name":                None,
            "file_info":                None,
            "node_type":                "",
        })
        for item in node_records
    ]

    nested_dicts = build_simplified_tree(node_records)

    def _dicts_to_tree_items(nodes_dict: List[Dict]) -> List:
        result = []
        for d in nodes_dict:
            item = TreeItem(
                eid=d.get("eid", ""),
                level=d.get("level", 1),
                idx=d.get("idx", 0),
                text=d.get("text", "") or d.get("title_text", ""),
                children=_dicts_to_tree_items(d.get("children", [])),
            )
            matched = next((n for n in tree_nodes_org if n.eid == item.eid), None)
            if matched:
                item.id                       = matched.id
                item.parent_id                = matched.parent_id
                item.file_path                = matched.file_path
                item.update_file_path         = matched.update_file_path
                item.is_conversion_completion = matched.is_conversion_completion
            result.append(item)
        return result

    nested_tree_items = _dicts_to_tree_items(nested_dicts)
    return process_split_tree_nodes_with_select(
        tree_nodes_org=nested_tree_items,
        record_id=rec_id,
        current_time=cur_time,
        file_base_path="",
    )


# ─── 嵌入组件（yxdl_embed_components）CRUD ────────────────────────────────────

def insert_embed_component(row: Dict[str, Any]) -> int:
    """插入一条嵌入组件记录，返回自增 id。row 由 spec_to_db_row() 产出。"""
    sql = """
        INSERT INTO "yxdl_embed_components"
            (embed_id, embed_type, version, title, display, url,
             payload, record_id, node_id)
        VALUES (%(embed_id)s, %(embed_type)s, %(version)s, %(title)s, %(display)s,
                %(url)s, %(payload)s::jsonb, %(record_id)s, %(node_id)s)
        RETURNING id
    """
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(sql, row)
            new_id = cursor.fetchone()[0]
            conn.commit()
    logger.info(f"insert_embed_component: embed_id={row['embed_id']} db_id={new_id}")
    return new_id


def get_embed_component(embed_id: str) -> Optional[Dict[str, Any]]:
    """按 embed_id 查询组件，不存在返回 None。"""
    sql = """
        SELECT id, embed_id, embed_type, version, title, display, url, payload,
               record_id, node_id, status, create_time, update_time
        FROM "yxdl_embed_components"
        WHERE embed_id = %s AND status = 1
    """
    with get_db_connection() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cursor:
            cursor.execute(sql, (embed_id,))
            row = cursor.fetchone()
    return dict(row) if row else None


def update_embed_component(embed_id: str, row: Dict[str, Any]) -> bool:
    """按 embed_id 更新组件；payload/title/url/version 可改，类型不可改。"""
    sql = """
        UPDATE "yxdl_embed_components"
        SET title = %(title)s,
            url = %(url)s,
            version = %(version)s,
            display = %(display)s,
            payload = %(payload)s::jsonb,
            update_time = CURRENT_TIMESTAMP
        WHERE embed_id = %(embed_id)s AND status = 1
    """
    params = {**row, "embed_id": embed_id}
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(sql, params)
            affected = cursor.rowcount
            conn.commit()
    return affected > 0


def delete_embed_component(embed_id: str) -> bool:
    """软删除（status=0）。"""
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(
                'UPDATE "yxdl_embed_components" SET status = 0, update_time = CURRENT_TIMESTAMP '
                'WHERE embed_id = %s AND status = 1',
                (embed_id,),
            )
            affected = cursor.rowcount
            conn.commit()
    return affected > 0


def list_embed_components_by_record(record_id: int) -> List[Dict[str, Any]]:
    """列出某个文档下所有组件。"""
    sql = """
        SELECT embed_id, embed_type, version, title, display, url, payload,
               record_id, node_id, create_time, update_time
        FROM "yxdl_embed_components"
        WHERE record_id = %s AND status = 1
        ORDER BY create_time ASC
    """
    with get_db_connection() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cursor:
            cursor.execute(sql, (record_id,))
            return [dict(r) for r in cursor.fetchall()]


def get_embed_components_by_ids(embed_ids: List[str]) -> Dict[str, Dict[str, Any]]:
    """批量按 embed_id 取组件，返回 {embed_id: row}。"""
    if not embed_ids:
        return {}
    sql = """
        SELECT embed_id, embed_type, version, title, display, url, payload,
               record_id, node_id
        FROM "yxdl_embed_components"
        WHERE embed_id = ANY(%s) AND status = 1
    """
    with get_db_connection() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cursor:
            cursor.execute(sql, (embed_ids,))
            return {row["embed_id"]: dict(row) for row in cursor.fetchall()}


def insert_xlsx_upload_record(
    original_filename: str,
    new_filename: str,
    file_sign: str,
    save_path: str,
) -> int:
    """向 yxdl_xlsx_upload_records 写入一条 xlsx 上传记录，返回新行 id。"""
    sql = """
        INSERT INTO "yxdl_xlsx_upload_records"
          (original_filename, new_filename, file_sign, save_path, upload_time, update_time)
        VALUES (%s, %s, %s, %s, NOW(), NOW())
        RETURNING id;
    """
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(sql, (original_filename, new_filename, file_sign, save_path))
            row_id = cursor.fetchone()[0]
            conn.commit()
    return row_id


def get_original_filename_by_new_filename(new_filename: str) -> Optional[str]:
    """按 new_filename 从 xlsx 上传记录表查 original_filename，未找到返回 None。"""
    sql = """
        SELECT original_filename FROM "yxdl_xlsx_upload_records"
        WHERE new_filename = %s
        LIMIT 1
    """
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(sql, (new_filename,))
            row = cursor.fetchone()
    return row[0] if row else None