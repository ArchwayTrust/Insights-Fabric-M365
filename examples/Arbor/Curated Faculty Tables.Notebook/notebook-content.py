# Fabric notebook source

# METADATA ********************

# META {
# META   "kernel_info": {
# META     "name": "synapse_pyspark"
# META   },
# META   "dependencies": {
# META     "lakehouse": {
# META       "default_lakehouse_name": "",
# META       "default_lakehouse_workspace_id": ""
# META     },
# META     "environment": {}
# META   }
# META }

# MARKDOWN ********************

# # Curated Faculty Tables
# Builds the faculty list and faculty membership tables from Arbor MIS data. The Manage Faculty Sharepoint Libraries
# notebook reads them to create each faculty's library and keep its members in line. The README next to this notebook
# says where each table it reads comes from.

# MARKDOWN ********************

# ## Setup

# CELL ********************

from notebookutils import variableLibrary

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# PARAMETERS CELL ********************

# Abort the merge when the source produces fewer rows than this fraction of the
# current live faculty list — protects against a failed/partial Arbor load
# soft-deleting every faculty (which would drive mass group-emptying downstream).
SOURCE_MIN_ROW_FRACTION = 0.8

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

# Get the entire variable library by name
vl = variableLibrary.getLibrary("VL_CORE")

# Fetch key environment variables
STORE_WORKSPACE_NAME = vl.getVariable("STORE_WORKSPACE_NAME")
BASE_LH_NAME         = vl.getVariable("BASE_LH_NAME")
CURATED_LH_NAME      = vl.getVariable("CURATED_LH_NAME")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

# Explicitly name all lakehouse tables. Education (edu_) and conformed (con_) tables are still in dbo
# under their domain prefix; IT tables have moved to the it schema.
base_staff                  = f"{STORE_WORKSPACE_NAME}.{BASE_LH_NAME}.dbo.edu_arbor_staff"
base_faculty_memberships    = f"{STORE_WORKSPACE_NAME}.{BASE_LH_NAME}.dbo.edu_arbor_staff_faculty_memberships"
base_department_memberships = f"{STORE_WORKSPACE_NAME}.{BASE_LH_NAME}.dbo.edu_arbor_staff_department_memberships"
base_departments            = f"{STORE_WORKSPACE_NAME}.{BASE_LH_NAME}.dbo.edu_arbor_departments"
base_faculties              = f"{STORE_WORKSPACE_NAME}.{BASE_LH_NAME}.dbo.edu_arbor_faculties"
base_faculty_departments    = f"{STORE_WORKSPACE_NAME}.{BASE_LH_NAME}.dbo.edu_arbor_faculty_departments"

curated_sites               = f"{STORE_WORKSPACE_NAME}.{CURATED_LH_NAME}.dbo.con_dim_sites"

curated_faculty_list        = f"{STORE_WORKSPACE_NAME}.{CURATED_LH_NAME}.it.sharepoint_faculty_list"
curated_faculty_membership  = f"{STORE_WORKSPACE_NAME}.{CURATED_LH_NAME}.it.sharepoint_faculty_membership"


print("Lakehouse tables:")
print(base_staff)
print(base_faculty_memberships)
print(base_department_memberships)
print(base_departments)
print(base_faculties)
print(base_faculty_departments)
print(curated_sites)
print(curated_faculty_list)
print(curated_faculty_membership)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# MARKDOWN ********************

# ## Create Temp Views

# CELL ********************

# Staff with a current, active faculty membership

df = spark.sql(f"""
    SELECT
        sfm.faculty_unique_id
        ,sfm.staff_unique_id
        ,s.legal_last_name
        ,s.legal_first_name
        ,s.email
    FROM
        {base_faculty_memberships} AS sfm
        LEFT JOIN {base_staff} AS s ON sfm.staff_unique_id = s.staff_unique_id
    WHERE
        (sfm.end_date IS NULL OR sfm.end_date >= CURRENT_DATE())
        AND sfm.is_active = 1
        AND s.has_active_roles = 1
""")

df.createOrReplaceTempView("tvw_faculty_staff")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

df = spark.sql(f"""
    SELECT
        sdm.department_unique_id
        ,sdm.staff_unique_id
        ,s.legal_last_name
        ,s.legal_first_name
        ,s.email
    FROM
        {base_department_memberships} AS sdm
        LEFT JOIN {base_staff} AS s ON sdm.staff_unique_id = s.staff_unique_id
    WHERE
        (sdm.end_date IS NULL OR sdm.end_date >= CURRENT_DATE())
        AND sdm.is_active = 1
        AND s.has_active_roles = 1
""")

df.createOrReplaceTempView("tvw_department_staff")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

# Departments joined to their owning faculty.  INNER JOINs state the intent
# (departments without a faculty mapping are irrelevant here); the date
# predicates live in the ON clause so a NULL start_date does not silently
# drop an otherwise-current mapping.

df = spark.sql(f"""
SELECT
    f.application_id
    ,d.department_unique_id
    ,f.faculty_unique_id
    ,sch.academy_code
    ,sch.sharepoint_hub_url
    ,d.department_name
    ,f.faculty_name
    ,f.faculty_code

FROM
    {base_faculties} AS f
    LEFT JOIN {base_faculty_departments} AS fd
        ON fd.faculty_unique_id = f.faculty_unique_id
        AND (fd.start_date IS NULL OR fd.start_date <= CURRENT_DATE())
        AND (fd.end_date   IS NULL OR fd.end_date   >= CURRENT_DATE())
    LEFT JOIN {base_departments} AS d
        ON d.department_unique_id = fd.department_unique_id
    LEFT JOIN {curated_sites} AS sch
        ON sch.arbor_application_id = COALESCE(d.application_id, f.application_id)
WHERE f.is_active = 1
""")

df.createOrReplaceTempView("tvw_department_faculties")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

df = spark.sql(f"""
    SELECT DISTINCT
        df.application_id
        ,df.faculty_unique_id
        ,df.academy_code
        ,df.sharepoint_hub_url
        ,df.faculty_name
        ,df.faculty_code
        ,ds.legal_last_name
        ,ds.legal_first_name
        ,ds.email AS member_email
    FROM
        tvw_department_faculties AS df
        LEFT JOIN tvw_department_staff AS ds ON ds.department_unique_id = df.department_unique_id
""")

df.createOrReplaceTempView("tvw_department_members")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

df = spark.sql(f"""
    SELECT
        df.application_id
        ,df.faculty_unique_id
        ,df.academy_code
        ,df.sharepoint_hub_url
        ,df.faculty_name
        ,df.faculty_code
        ,fs.legal_last_name
        ,fs.legal_first_name
        ,fs.email AS member_email
    FROM
        tvw_department_faculties AS df
        LEFT JOIN tvw_faculty_staff AS fs ON df.faculty_unique_id = fs.faculty_unique_id
""")

df.createOrReplaceTempView("tvw_faculty_members")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

df = spark.sql(f"""
    SELECT *
    FROM tvw_department_members

    UNION

    SELECT *
    FROM tvw_faculty_members
""")

df.createOrReplaceTempView("tvw_member_union")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# MARKDOWN ********************

# ## Merge Faculty Table
# This maintains an SCD-style table (see `curated_faculty_list` defined in Setup) used to
# track changes and inform sharepoint library updates.
# Change signalling uses `last_changed_ts`: every clause that represents an event the
# consumer must act on (insert, rename, resurrect, soft delete) stamps it with the merge
# timestamp. The consumer (the Manage Faculty Sharepoint Libraries notebook) processes rows newer than its own
# stored watermark, so events survive repeated producer runs. Pure membership-count
# changes do NOT bump the watermark — membership sync is a full reconcile each run.

# CELL ********************

# COUNT(member_email) rather than COUNT(1): the union contains a NULL-email
# placeholder row for faculties with no current members, which must count as 0.
# faculty_unique_id is the merge key, so NULL keys are excluded outright.

df = spark.sql(f"""
    SELECT
        m.application_id
        ,m.faculty_unique_id
        ,m.academy_code
        ,m.sharepoint_hub_url
        ,m.faculty_name
        ,m.faculty_code
        ,COUNT(m.member_email) AS membership_count
    FROM tvw_member_union AS m
    WHERE m.faculty_unique_id IS NOT NULL
    GROUP BY ALL
""")

df.createOrReplaceTempView("tvw_faculty_source")

# The view sits on a multi-table join that the two sanity guards and the
# MERGE would otherwise each recompute from scratch — cache it once for all
# three (released again straight after the merge).
spark.sql("CACHE TABLE tvw_faculty_source")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

# Ensure the faculty list table exists
# CREATE TABLE does not create its schema — a fresh schema-enabled
# lakehouse fails with SCHEMA_NOT_FOUND without this.
spark.sql(f"CREATE SCHEMA IF NOT EXISTS {STORE_WORKSPACE_NAME}.{CURATED_LH_NAME}.it")

spark.sql(f"""
    CREATE TABLE IF NOT EXISTS {curated_faculty_list} (
        application_id          STRING,
        faculty_unique_id       STRING,
        academy_code            STRING,
        sharepoint_hub_url      STRING,
        previous_hub_url        STRING,     -- Old hub, recorded when the hub URL changes; consumer retires the old-site library.
        current_faculty_name    STRING,
        static_faculty_code     STRING,
        created_date            TIMESTAMP,
        previous_faculty_name   STRING,
        -- previous_faculty_code   STRING,  -- Faculty_code is used for the sharepoint path, safer if this remains constant and code only updates display name.
        date_name_changed       TIMESTAMP,
        membership_count        BIGINT,
        deleted_in_source       BOOLEAN,
        last_changed_ts         TIMESTAMP,  -- Consumer watermark: stamped by every clause the consumer must act on.
        exclude_from_sharepoint BOOLEAN,    -- Manually set flag if a faculty shouldn't have a library created.
                                            -- When flipping it by hand, also set last_changed_ts = current_timestamp():
                                            -- the consumer is watermark-gated and never revisits rows behind it.
        -- Manually controlled all-staff access. The merge only sets the defaults on insert and never
        -- updates them. Membership sync is a full pass, so changes apply next run with no
        -- last_changed_ts bump.
        all_staff_read           INT,       -- 1 = nest all_staff_security_group in the library's Auto-Read group.
        all_staff_contribute     INT,       -- 1 = nest all_staff_security_group in the library's Auto-Contribute group.
        all_staff_security_group STRING     -- Entra group to nest: exact display name or object ID.
    )
    USING DELTA
""")

# Migrate tables created before the watermark redesign: add last_changed_ts and
# backfill it. The legacy changed_today column (if present) is left in place but
# is no longer read or written.
existing_cols = {f.name for f in spark.table(curated_faculty_list).schema.fields}
if "last_changed_ts" not in existing_cols:
    spark.sql(f"ALTER TABLE {curated_faculty_list} ADD COLUMNS (last_changed_ts TIMESTAMP)")
    spark.sql(f"""
        UPDATE {curated_faculty_list}
        SET last_changed_ts = COALESCE(date_name_changed, created_date, current_timestamp())
    """)
    print("Migrated: added last_changed_ts and backfilled from date_name_changed/created_date.")
if "previous_hub_url" not in existing_cols:
    spark.sql(f"ALTER TABLE {curated_faculty_list} ADD COLUMNS (previous_hub_url STRING)")
    print("Migrated: added previous_hub_url.")
# ADD COLUMNS leaves existing rows NULL — backfill the flags to their 0 default
# so the table reads the same whether a row predates the columns or not.
for flag_col in ("all_staff_read", "all_staff_contribute"):
    if flag_col not in existing_cols:
        spark.sql(f"ALTER TABLE {curated_faculty_list} ADD COLUMNS ({flag_col} INT)")
        spark.sql(f"UPDATE {curated_faculty_list} SET {flag_col} = 0")
        print(f"Migrated: added {flag_col} and backfilled 0.")
if "all_staff_security_group" not in existing_cols:
    spark.sql(f"ALTER TABLE {curated_faculty_list} ADD COLUMNS (all_staff_security_group STRING)")
    print("Migrated: added all_staff_security_group.")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

# Source sanity guards — abort BEFORE the merge rather than cascade a bad load
# into mass soft-deletes (which would drive the consumer to empty every group).

source_count = spark.sql("SELECT COUNT(*) AS c FROM tvw_faculty_source").collect()[0]["c"]

duplicate_keys = spark.sql("""
    SELECT faculty_unique_id, COUNT(*) AS c
    FROM tvw_faculty_source
    GROUP BY faculty_unique_id
    HAVING COUNT(*) > 1
""").collect()
if duplicate_keys:
    sample = ", ".join(f"{r['faculty_unique_id']} (x{r['c']})" for r in duplicate_keys[:10])
    raise ValueError(
        f"tvw_faculty_source has {len(duplicate_keys)} duplicated faculty_unique_id value(s) — "
        f"merge would be ambiguous. Likely a dim_sites or faculty_departments fan-out. Sample: {sample}"
    )

live_count = spark.sql(f"""
    SELECT COUNT(*) AS c FROM {curated_faculty_list} WHERE deleted_in_source = false
""").collect()[0]["c"]

if live_count > 0:
    if source_count == 0:
        raise ValueError(
            f"tvw_faculty_source is empty but {curated_faculty_list} has {live_count} live rows — "
            f"refusing to soft-delete everything. Check the upstream Arbor load."
        )
    if source_count < SOURCE_MIN_ROW_FRACTION * live_count:
        raise ValueError(
            f"tvw_faculty_source has {source_count} rows, below {SOURCE_MIN_ROW_FRACTION:.0%} of the "
            f"{live_count} live rows in {curated_faculty_list} — refusing to merge a suspiciously "
            f"small source. Override by lowering SOURCE_MIN_ROW_FRACTION if the shrink is genuine."
        )

print(f"Source sanity OK: {source_count} source rows vs {live_count} live rows.")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# CELL ********************

result = spark.sql(f"""
    MERGE INTO {curated_faculty_list} AS t
    USING tvw_faculty_source AS s
    ON t.faculty_unique_id  = s.faculty_unique_id

    -- name changed → shift current into previous, stamp the watermark
    -- (also resurrects if the row was soft-deleted: deleted_in_source is reset;
    --  academy/hub refreshed so a late-arriving dim_sites row is picked up)
    -- Every clause that refreshes the hub URL also records the old hub in
    -- previous_hub_url (only when it actually changed) so the consumer can
    -- retire the library left behind on the old site.  MERGE SET clauses all
    -- read the pre-update row, so t.sharepoint_hub_url is the old value.
    WHEN MATCHED
        AND t.current_faculty_name IS DISTINCT FROM s.faculty_name
    THEN UPDATE SET
        t.previous_faculty_name = t.current_faculty_name,
        t.current_faculty_name  = s.faculty_name,
        t.academy_code          = s.academy_code,
        t.previous_hub_url      = CASE WHEN t.sharepoint_hub_url IS DISTINCT FROM s.sharepoint_hub_url
                                       THEN t.sharepoint_hub_url ELSE t.previous_hub_url END,
        t.sharepoint_hub_url    = s.sharepoint_hub_url,
        t.membership_count      = s.membership_count,
        t.date_name_changed     = current_timestamp(),
        t.deleted_in_source     = false,
        t.last_changed_ts       = current_timestamp()

    -- was soft-deleted, now reappeared → resurrect, stamp the watermark
    WHEN MATCHED
        AND t.deleted_in_source = true
    THEN UPDATE SET
        t.academy_code          = s.academy_code,
        t.previous_hub_url      = CASE WHEN t.sharepoint_hub_url IS DISTINCT FROM s.sharepoint_hub_url
                                       THEN t.sharepoint_hub_url ELSE t.previous_hub_url END,
        t.sharepoint_hub_url    = s.sharepoint_hub_url,
        t.membership_count      = s.membership_count,
        t.deleted_in_source     = false,
        t.last_changed_ts       = current_timestamp()

    -- academy/hub changed (e.g. dim_sites populated after first insert, or a
    -- site migration) → refresh and stamp the watermark so the consumer acts
    WHEN MATCHED
        AND (t.academy_code IS DISTINCT FROM s.academy_code
             OR t.sharepoint_hub_url IS DISTINCT FROM s.sharepoint_hub_url)
    THEN UPDATE SET
        t.academy_code          = s.academy_code,
        t.previous_hub_url      = CASE WHEN t.sharepoint_hub_url IS DISTINCT FROM s.sharepoint_hub_url
                                       THEN t.sharepoint_hub_url ELSE t.previous_hub_url END,
        t.sharepoint_hub_url    = s.sharepoint_hub_url,
        t.membership_count      = s.membership_count,
        t.last_changed_ts       = current_timestamp()

    -- only membership change → update the count; NOT a watermark event
    -- (membership sync is a full reconcile, it does not consume the watermark)
    WHEN MATCHED
        AND t.membership_count IS DISTINCT FROM s.membership_count
    THEN UPDATE SET
        t.membership_count      = s.membership_count

    -- new row (the manually controlled columns — exclude_from_sharepoint and
    -- all_staff_* — are only defaulted here; no MATCHED clause touches them)
    WHEN NOT MATCHED
    THEN INSERT (
        application_id,
        faculty_unique_id,
        academy_code,
        sharepoint_hub_url,
        previous_hub_url,
        current_faculty_name,
        static_faculty_code,
        created_date,
        previous_faculty_name,
        date_name_changed,
        membership_count,
        deleted_in_source,
        last_changed_ts,
        exclude_from_sharepoint,
        all_staff_read,
        all_staff_contribute,
        all_staff_security_group
    )
    VALUES (
        s.application_id,
        s.faculty_unique_id,
        s.academy_code,
        s.sharepoint_hub_url,
        NULL,
        s.faculty_name,
        s.faculty_code,
        current_timestamp(),
        NULL,
        NULL,
        s.membership_count,
        false,
        current_timestamp(),
        false,
        0,
        0,
        NULL
    )

    -- gone from source → soft delete once; the watermark stamp makes the
    -- consumer unlink it exactly once (no re-unlink on later runs)
    WHEN NOT MATCHED BY SOURCE
        AND t.deleted_in_source = false
    THEN UPDATE SET
        t.deleted_in_source = true,
        t.last_changed_ts   = current_timestamp(),
        t.membership_count  = 0
""")

row = result.collect()[0]

print(f"Merge complete: {row.num_inserted_rows} inserted, {row.num_updated_rows} updated, {row.num_deleted_rows} deleted ({row.num_affected_rows} total affected)")

spark.sql("UNCACHE TABLE IF EXISTS tvw_faculty_source")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }

# MARKDOWN ********************

# ## Write Current Memberships
# Stores the current membership for each faculty in `curated_faculty_membership` (see Setup).
# No historic tracking required here but we do look up the original faculty_code as it was
# when the site was created from the Faculties SCD above. Faculties with zero current
# members simply have no rows here — the consumer drives its library list from the faculty
# list table and treats a missing membership set as "desired members: none".

# CELL ********************

df = spark.sql(f"""
    SELECT
        m.application_id
        ,m.faculty_unique_id
        ,m.academy_code
        ,m.sharepoint_hub_url
        ,m.faculty_name
        ,f.static_faculty_code
        ,m.legal_last_name
        ,m.legal_first_name
        ,m.member_email
    FROM tvw_member_union AS m
    LEFT JOIN {curated_faculty_list} AS f ON f.faculty_unique_id = m.faculty_unique_id
    WHERE m.member_email IS NOT NULL
""")

(df.write
    .format("delta")
    .mode("overwrite")
    .saveAsTable(curated_faculty_membership)
)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "synapse_pyspark"
# META }
