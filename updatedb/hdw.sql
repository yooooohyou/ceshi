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


CREATE TABLE "yxdl_docx_title_trees" (
  "id" SERIAL PRIMARY KEY,
  "record_id" int4,
  "title_text" varchar(512) COLLATE "pg_catalog"."default" NOT NULL,
  "html_content" text COLLATE "pg_catalog"."default",
  "create_time" timestamptz NOT NULL,
  "update_time" timestamptz NOT NULL DEFAULT CURRENT_TIMESTAMP,
  "level" int4,
  "eid" varchar(128) COLLATE "pg_catalog"."default",
  "idx" int4,
  "node_type" varchar(16) COLLATE "pg_catalog"."default" DEFAULT 'main',
  "origin_file_path" text COLLATE "pg_catalog"."default",
  "is_conversion_completion" int4,
  "update_file_path" text COLLATE "pg_catalog"."default",
  "parent_id" int4,
  "batch_count" int4,
  "split_id" int4
);

COMMENT ON COLUMN "yxdl_docx_title_trees"."id" IS '节点ID';
COMMENT ON COLUMN "yxdl_docx_title_trees"."record_id" IS '关联文件记录ID';
COMMENT ON COLUMN "yxdl_docx_title_trees"."title_text" IS '标题文本';
COMMENT ON COLUMN "yxdl_docx_title_trees"."html_content" IS 'Word转换后的HTML文本';
COMMENT ON COLUMN "yxdl_docx_title_trees"."create_time" IS '创建时间';
COMMENT ON COLUMN "yxdl_docx_title_trees"."update_time" IS '更新时间';
COMMENT ON COLUMN "yxdl_docx_title_trees"."level" IS '节点层级';
COMMENT ON COLUMN "yxdl_docx_title_trees"."eid" IS '拆分接口返回的eid';
COMMENT ON COLUMN "yxdl_docx_title_trees"."idx" IS '拆分接口返回的idx';
COMMENT ON COLUMN "yxdl_docx_title_trees"."node_type" IS '节点类型：main/branch';
COMMENT ON COLUMN "yxdl_docx_title_trees"."origin_file_path" IS '原始服务器文件路径';
COMMENT ON COLUMN "yxdl_docx_title_trees"."is_conversion_completion" IS '是否已经转换';
COMMENT ON COLUMN "yxdl_docx_title_trees"."update_file_path" IS '更新后的文件路径';
COMMENT ON COLUMN "yxdl_docx_title_trees"."parent_id" IS '父节点ID，NULL表示根节点';
COMMENT ON COLUMN "yxdl_docx_title_trees"."batch_count" IS '批次ID，第一次导入为1，每次更新在原最大值基础上+1';
COMMENT ON COLUMN "yxdl_docx_title_trees"."split_id" IS '拆分接口返回的树节点id，用于合并时还原树结构';
COMMENT ON TABLE "yxdl_docx_title_trees" IS '标题树节点表';


CREATE TABLE "yxdl_docx_upload_records" (
  "id" SERIAL PRIMARY KEY,
  "original_filename" varchar(255) COLLATE "pg_catalog"."default" NOT NULL,
  "new_filename" varchar(255) COLLATE "pg_catalog"."default" NOT NULL,
  "save_path" varchar(512) COLLATE "pg_catalog"."default" NOT NULL,
  "upload_time" timestamptz NOT NULL,
  "update_time" timestamptz NOT NULL DEFAULT CURRENT_TIMESTAMP,
  "split_file_id" varchar(128) COLLATE "pg_catalog"."default",
  "process_mode" varchar(16) COLLATE "pg_catalog"."default" DEFAULT 'single',
  "title_font_dict" jsonb DEFAULT NULL
);

COMMENT ON COLUMN "yxdl_docx_upload_records"."id" IS '记录ID';
COMMENT ON COLUMN "yxdl_docx_upload_records"."original_filename" IS '原始文件名';
COMMENT ON COLUMN "yxdl_docx_upload_records"."new_filename" IS '新文件名';
COMMENT ON COLUMN "yxdl_docx_upload_records"."save_path" IS '文件保存路径';
COMMENT ON COLUMN "yxdl_docx_upload_records"."upload_time" IS '上传时间';
COMMENT ON COLUMN "yxdl_docx_upload_records"."update_time" IS '更新时间';
COMMENT ON COLUMN "yxdl_docx_upload_records"."split_file_id" IS '拆分接口使用的file_id';
COMMENT ON COLUMN "yxdl_docx_upload_records"."process_mode" IS '处理模式：single/split';
COMMENT ON COLUMN "yxdl_docx_upload_records"."title_font_dict" IS '拆分接口返回的标题字体字典';
COMMENT ON TABLE "yxdl_docx_upload_records" IS 'DOCX文件上传记录';