import pandas as pd
from sqlalchemy import create_engine
import os
import traceback
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

# Dashboard JSON layer (serialisation only -- no KPI logic)
import merch_kpi_export

# ==========================================
# PostgreSQL Connection
# ==========================================

import os

load_dotenv(Path(__file__).resolve().parent / ".env")

DB_URL = os.getenv("DATABASE_URL")
if not DB_URL:
    raise ValueError("DATABASE_URL is missing. Check your .env file.")

try:
    engine = create_engine(DB_URL)

    with engine.connect() as conn:
        print("✅ Connected to PostgreSQL Successfully")

except Exception as e:
    print("❌ Database Connection Error")
    traceback.print_exc()
    exit()

# ==========================================
# Base Folder
# ==========================================

base_folder = r"G:\.shortcut-targets-by-id\1-OzYqibxMbDT38omFZxDLknsTKiBDbpt\Business Analytics\Dashboard\Supporting_Folder_For_Complete_dashboard\Merchandising_KPI_Dashboard"
if not os.path.exists(base_folder):
    print(f"❌ Base folder not found:\n{base_folder}")
    exit()

# Run timestamp used in the Excel file names
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

# ==========================================
# Reports Configuration
# ==========================================

reports = [
    {
        "folder": "mtd_sku_sale_vs_soh",
        "file": "mtd_sku_sale_vs_soh_14_oct.xlsx",
        "query": """
        WITH mtd_bounds AS (

    /* =========================================================
    MTD DATE WINDOW (DYNAMIC - NEVER HARDCODED)
    1st of current month -> yesterday
    ========================================================= */

    SELECT
        date_trunc('month', CURRENT_DATE)::date AS mtd_start,
        (CURRENT_DATE - INTERVAL '1 day')::date AS mtd_end
),

/* =========================================================
CTE A: SALES CTE
MTD sales aggregated by STORE + PRODUCT_CODE
========================================================= */

sales_cte AS (

    SELECT

        COALESCE(
            scm.store_code,
            TRIM(s.store_code::text)
        ) AS store_code,

        UPPER(
            TRIM(s.product_code::text)
        ) AS product_code,

        MAX(s.product_name) AS product_name,
        MAX(s.category) AS category,
        MAX(s.sub_class) AS sub_class,
        MAX(s.barcode) AS barcode,

        SUM(
            COALESCE(s.net_sales_qty,0)
        ) AS mtd_sale_qty,

        SUM(
            COALESCE(s.net_sales,0)
        ) AS mtd_sale_value

    FROM last_3_month_sale s

    LEFT JOIN store_code_mapping scm
        ON TRIM(s.store_code::text)
         = TRIM(scm.old_store_code::text)

    CROSS JOIN mtd_bounds mb

    WHERE safe_to_date(s.date::text) IS NOT NULL
      AND safe_to_date(s.date::text)
          BETWEEN mb.mtd_start AND mb.mtd_end

    GROUP BY 1,2
),

/* =========================================================
LATEST STOCK SNAPSHOT DATE (UP TO YESTERDAY)
Same snapshot-selection logic used in grid_by_count
========================================================= */

stock_snapshot_date AS (

    SELECT
        MAX(
            TO_TIMESTAMP(
                date,
                'YYYY-MM-DD HH24:MI:SS'
            )::date
        ) AS snapshot_dt

    FROM inv_90_days

    CROSS JOIN mtd_bounds mb

    WHERE
        TO_TIMESTAMP(
            date,
            'YYYY-MM-DD HH24:MI:SS'
        )::date <= mb.mtd_end
),

/* =========================================================
CTE B: STOCK CTE
Latest applicable SOH by STORE + PRODUCT_CODE
========================================================= */

stock_cte AS (

    SELECT

        COALESCE(
            scm.store_code,
            i.store_code
        ) AS store_code,

        TRIM(i.product_code::text) AS product_code,

        MAX(i.product_name) AS product_name,
        MAX(i.category) AS category,
        MAX(i.sub_class_name) AS sub_class,
        MAX(i.barcode) AS barcode,
        MAX(i.price) AS mrp,

        SUM(
            COALESCE(i.available_stock,0)
        ) AS available_stock,

        SUM(
            COALESCE(i.total_stock,0)
        ) AS total_stock

    FROM inv_90_days i

    LEFT JOIN store_code_mapping scm
        ON TRIM(i.store_code::text)
         = TRIM(scm.old_store_code::text)

    CROSS JOIN stock_snapshot_date sd

    WHERE
        TO_TIMESTAMP(
            i.date,
            'YYYY-MM-DD HH24:MI:SS'
        )::date = sd.snapshot_dt

    GROUP BY 1,2
)

/* =========================================================
FINAL RESULT
GRAIN: ONE ROW PER STORE + PRODUCT_CODE
========================================================= */

SELECT

    COALESCE(sc.store_code, st.store_code) AS store,
    COALESCE(sc.product_code, st.product_code) AS product_code,

    COALESCE(sc.product_name, st.product_name) AS product_name,
    COALESCE(sc.category, st.category) AS category,
    COALESCE(sc.sub_class, st.sub_class) AS sub_class,
    COALESCE(sc.barcode, st.barcode) AS barcode,
    st.mrp,

    COALESCE(sc.mtd_sale_qty,0) AS mtd_sale_qty,
    COALESCE(sc.mtd_sale_value,0) AS mtd_sale_value,

    st.available_stock AS available,
    st.total_stock

FROM sales_cte sc

FULL OUTER JOIN stock_cte st
    ON sc.store_code = st.store_code
   AND sc.product_code = st.product_code

ORDER BY
    1,2;
        """
    },
    {
        "folder": "new_sku_sales",
        "file": "new_sku_sales_14_oct.xlsx",
        "query": """
        WITH date_bounds AS (

    SELECT

        /* LAST 7 COMPLETED WEEKS */

        TO_CHAR(CURRENT_DATE - INTERVAL '1 week', 'IYYYIW')::bigint AS w1,
        TO_CHAR(CURRENT_DATE - INTERVAL '2 week', 'IYYYIW')::bigint AS w2,
        TO_CHAR(CURRENT_DATE - INTERVAL '3 week', 'IYYYIW')::bigint AS w3,
        TO_CHAR(CURRENT_DATE - INTERVAL '4 week', 'IYYYIW')::bigint AS w4,
        TO_CHAR(CURRENT_DATE - INTERVAL '5 week', 'IYYYIW')::bigint AS w5,
        TO_CHAR(CURRENT_DATE - INTERVAL '6 week', 'IYYYIW')::bigint AS w6,
        TO_CHAR(CURRENT_DATE - INTERVAL '7 week', 'IYYYIW')::bigint AS w7,

        date_trunc('month', CURRENT_DATE - INTERVAL '1 month')::date AS m1,
        date_trunc('month', CURRENT_DATE - INTERVAL '2 month')::date AS m2,
        date_trunc('month', CURRENT_DATE - INTERVAL '3 month')::date AS m3,
        date_trunc('month', CURRENT_DATE - INTERVAL '4 month')::date AS m4
),

/* =========================================================
STORE CODE NORMALIZATION
========================================================= */

store_code_normalized AS (

    SELECT DISTINCT
        TRIM(old_store_code::text) AS source_code,
        TRIM(store_code::text)     AS final_store_code
    FROM store_code_mapping

    UNION

    SELECT DISTINCT
        TRIM(store_code::text)     AS source_code,
        TRIM(store_code::text)     AS final_store_code
    FROM store_code_mapping
),

/* =========================================================
SALES BASE
========================================================= */

sales_base AS (

    SELECT

        COALESCE(
            scn.final_store_code,
            TRIM(s.store_code::text)
        ) AS store_code,

        s.date::timestamp::date AS sales_date,

        TO_CHAR(
            s.date::timestamp::date,
            'IYYYIW'
        )::bigint AS week_id,

        date_trunc(
            'month',
            s.date::timestamp::date
        )::date AS month_start,

        UPPER(
            TRIM(s.product_code::text)
        ) AS product_code,

        SUM(
            COALESCE(s.net_sales,0)
        ) AS net_sales_value,

        SUM(
            COALESCE(s.net_sales_qty,0)
        ) AS net_sales_qty

    FROM last_3_month_sale s

    LEFT JOIN store_code_normalized scn
        ON TRIM(s.store_code::text)
        = scn.source_code

    WHERE safe_to_date(s.date::text) IS NOT NULL
      AND safe_to_date(s.date::text) >= CURRENT_DATE - INTERVAL '180 day'

    GROUP BY
        1,2,3,4,5
),

/* =========================================================
NEW SKU MASTER
========================================================= */

new_sku_master AS (

    SELECT DISTINCT

        UPPER(
            TRIM(product_code::text)
        ) AS product_code

    FROM product_master

    WHERE b2b_launch_date IS NOT NULL

      AND safe_to_date(b2b_launch_date::text) IS NOT NULL

      AND safe_to_date(b2b_launch_date::text) >= CURRENT_DATE - INTERVAL '30 day'
),

/* =========================================================
NEW SKU SALES
========================================================= */

new_sku_sales AS (

    SELECT

        sb.store_code,
        sb.sales_date,
        sb.week_id,
        sb.month_start,

        SUM(sb.net_sales_value) AS new_sku_sales_value,

        SUM(sb.net_sales_qty) AS new_sku_sales_qty

    FROM sales_base sb

    INNER JOIN new_sku_master nm
        ON sb.product_code = nm.product_code

    GROUP BY
        sb.store_code,
        sb.sales_date,
        sb.week_id,
        sb.month_start
),

/* =========================================================
TOTAL STORE SALES
========================================================= */

total_store_sales AS (

    SELECT

        store_code,
        sales_date,
        week_id,
        month_start,

        SUM(net_sales_value) AS total_store_sales_value,

        SUM(net_sales_qty) AS total_store_sales_qty

    FROM sales_base

    GROUP BY
        store_code,
        sales_date,
        week_id,
        month_start
),

/* =========================================================
DAILY KPI
========================================================= */

kpi_daily AS (

    SELECT

        ts.store_code,
        ts.sales_date,
        ts.week_id,
        ts.month_start,

        COALESCE(

            ROUND(

                (
                    COALESCE(
                        ns.new_sku_sales_value,
                        0
                    )
                    /
                    NULLIF(
                        ts.total_store_sales_value,
                        0
                    )
                ) * 100,

                2

            ),

            0

        ) AS sales_contribution_pct,

        COALESCE(

            ROUND(

                (
                    COALESCE(
                        ns.new_sku_sales_qty,
                        0
                    )
                    /
                    NULLIF(
                        ts.total_store_sales_qty,
                        0
                    )
                ) * 100,

                2

            ),

            0

        ) AS qty_contribution_pct

    FROM total_store_sales ts

    LEFT JOIN new_sku_sales ns

        ON ts.store_code = ns.store_code
       AND ts.sales_date = ns.sales_date
),

/* =========================================================
WEEKLY KPI
========================================================= */

weekly_kpi AS (

    SELECT

        store_code,
        week_id,

        ROUND(
            AVG(sales_contribution_pct),
            2
        ) AS weekly_sales_contribution_pct

    FROM kpi_daily

    GROUP BY
        store_code,
        week_id
),

/* =========================================================
MONTHLY KPI
========================================================= */

monthly_kpi AS (

    SELECT

        store_code,
        month_start,

        ROUND(
            AVG(sales_contribution_pct),
            2
        ) AS monthly_sales_contribution_pct

    FROM kpi_daily

    GROUP BY
        store_code,
        month_start
)

/* =========================================================
FINAL OUTPUT
========================================================= */

SELECT

    wk.store_code,

    COALESCE(MAX(CASE WHEN wk.week_id = d.w1 THEN wk.weekly_sales_contribution_pct END),0) AS w1_sales_contribution_pct,
    COALESCE(MAX(CASE WHEN wk.week_id = d.w2 THEN wk.weekly_sales_contribution_pct END),0) AS w2_sales_contribution_pct,
    COALESCE(MAX(CASE WHEN wk.week_id = d.w3 THEN wk.weekly_sales_contribution_pct END),0) AS w3_sales_contribution_pct,
    COALESCE(MAX(CASE WHEN wk.week_id = d.w4 THEN wk.weekly_sales_contribution_pct END),0) AS w4_sales_contribution_pct,
    COALESCE(MAX(CASE WHEN wk.week_id = d.w5 THEN wk.weekly_sales_contribution_pct END),0) AS w5_sales_contribution_pct,
    COALESCE(MAX(CASE WHEN wk.week_id = d.w6 THEN wk.weekly_sales_contribution_pct END),0) AS w6_sales_contribution_pct,
    COALESCE(MAX(CASE WHEN wk.week_id = d.w7 THEN wk.weekly_sales_contribution_pct END),0) AS w7_sales_contribution_pct,

    COALESCE(MAX(CASE WHEN mk.month_start = d.m1 THEN mk.monthly_sales_contribution_pct END),0) AS m1_sales_contribution_pct,
    COALESCE(MAX(CASE WHEN mk.month_start = d.m2 THEN mk.monthly_sales_contribution_pct END),0) AS m2_sales_contribution_pct,
    COALESCE(MAX(CASE WHEN mk.month_start = d.m3 THEN mk.monthly_sales_contribution_pct END),0) AS m3_sales_contribution_pct,
    COALESCE(MAX(CASE WHEN mk.month_start = d.m4 THEN mk.monthly_sales_contribution_pct END),0) AS m4_sales_contribution_pct

FROM weekly_kpi wk

LEFT JOIN monthly_kpi mk
       ON wk.store_code = mk.store_code

CROSS JOIN date_bounds d

GROUP BY
    wk.store_code

ORDER BY
    wk.store_code;


        """
     },
     {
        "folder": "moh",
        "file": "moh_14_oct.xlsx",
        "query": """
        WITH date_bounds AS (

    SELECT

        CURRENT_DATE - INTERVAL '125 day' AS start_date,

        CURRENT_DATE - INTERVAL '1 day' AS d1,
        CURRENT_DATE - INTERVAL '2 day' AS d2,
        CURRENT_DATE - INTERVAL '3 day' AS d3,
        CURRENT_DATE - INTERVAL '4 day' AS d4,
        CURRENT_DATE - INTERVAL '5 day' AS d5,
        CURRENT_DATE - INTERVAL '6 day' AS d6,
        CURRENT_DATE - INTERVAL '7 day' AS d7,

        /* WEEKLY SNAPSHOT DATES */

(
    CURRENT_DATE
    - EXTRACT(DOW FROM CURRENT_DATE)::int
) AS w1,

(
    CURRENT_DATE
    - EXTRACT(DOW FROM CURRENT_DATE)::int
    - INTERVAL '7 day'
) AS w2,

(
    CURRENT_DATE
    - EXTRACT(DOW FROM CURRENT_DATE)::int
    - INTERVAL '14 day'
) AS w3,

(
    CURRENT_DATE
    - EXTRACT(DOW FROM CURRENT_DATE)::int
    - INTERVAL '21 day'
) AS w4,

/* MONTHLY SNAPSHOT DATES */

(
    DATE_TRUNC('month', CURRENT_DATE)
    - INTERVAL '1 day'
)::date AS m1,

(
    DATE_TRUNC('month', CURRENT_DATE)
    - INTERVAL '1 month'
    - INTERVAL '1 day'
)::date AS m2,

(
    DATE_TRUNC('month', CURRENT_DATE)
    - INTERVAL '2 month'
    - INTERVAL '1 day'
)::date AS m3

),

/* =====================================================
STORE MAP
===================================================== */

store_map AS (

    SELECT

        TRIM(old_store_code::text) AS old_store_code,

        TRIM(store_code::text) AS new_store_code

    FROM store_code_mapping

),

/* =====================================================
STORE NAMES
===================================================== */

store_names AS (

    SELECT

        COALESCE(

            sm.new_store_code,

            TRIM(i.store_code::text)

        ) AS store_code,

        MAX(i.store_name) AS store_name

    FROM inv_90_days i

    LEFT JOIN store_map sm

        ON TRIM(i.store_code::text)
        = sm.old_store_code

    GROUP BY 1

),

/* =====================================================
SALES DAILY
===================================================== */

sales_daily AS (

    SELECT

        COALESCE(

            sm.new_store_code,

            TRIM(s.store_code::text)

        ) AS store_code,

        CAST(LEFT(s.date,10) AS DATE) AS dt,

        /* Keep negative sales */

        SUM(

            COALESCE(s.net_sales_qty,0)

        ) AS daily_sales

    FROM (
        SELECT *
        FROM last_3_month_sale
        WHERE date::text ~ '^\d{4}-\d{2}-\d{2}'
    ) s

    LEFT JOIN store_map sm

        ON TRIM(s.store_code::text)
        = sm.old_store_code

    CROSS JOIN date_bounds db

    WHERE CAST(LEFT(s.date,10) AS DATE)
          >= db.start_date

      AND COALESCE(s.category,'')

          NOT IN (

              'Consumable',

              'Consumables',

              'Decoration Materials'

          )

    GROUP BY

        1,

        2

),
/* =====================================================
STOCK DAILY
===================================================== */

stock_daily AS (

    SELECT

        COALESCE(

            sm.new_store_code,

            TRIM(i.store_code::text)

        ) AS store_code,

        CAST(LEFT(i.date,10) AS DATE) AS dt,

        /* Ignore negative inventory */

        SUM(

            GREATEST(

                COALESCE(i.total_stock,0),

                0

            )

        ) AS daily_stock

    FROM (
        SELECT *
        FROM inv_90_days
        WHERE date::text ~ '^\d{4}-\d{2}-\d{2}'
    ) i

    LEFT JOIN store_map sm

        ON TRIM(i.store_code::text)
        = sm.old_store_code

    CROSS JOIN date_bounds db

    WHERE CAST(LEFT(i.date,10) AS DATE)
          >= db.start_date

      AND COALESCE(i.category,'')

          NOT IN (

              'Consumable',

              'Consumables',

              'Decoration Materials'

          )

    GROUP BY

        1,

        2

),

/* =====================================================
ALL STORE-DATE COMBINATIONS
===================================================== */

all_dates AS (

    SELECT DISTINCT

        store_code,

        dt

    FROM sales_daily

    UNION

    SELECT DISTINCT

        store_code,

        dt

    FROM stock_daily

),

/* =====================================================
BASE
===================================================== */

base AS (

    SELECT

        a.store_code,

        a.dt,

        COALESCE(

            st.daily_stock,

            0

        ) AS daily_stock,

        COALESCE(

            sa.daily_sales,

            0

        ) AS daily_sales,

        TO_CHAR(

            a.dt,

            'IYYYIW'

        )::bigint AS week_id,

        TO_CHAR(

            a.dt,

            'YYYY-MM'

        ) AS month_id

    FROM all_dates a

    LEFT JOIN stock_daily st

        ON a.store_code = st.store_code

       AND a.dt = st.dt

    LEFT JOIN sales_daily sa

        ON a.store_code = sa.store_code

       AND a.dt = sa.dt

),

/* =====================================================
ROLLING 30 DAY SALES
===================================================== */

calc AS (

    SELECT

        store_code,

        dt,

        week_id,

        month_id,

        daily_stock,

        daily_sales,

        SUM(daily_sales) OVER (

            PARTITION BY store_code

            ORDER BY dt

            ROWS BETWEEN 29 PRECEDING
                     AND CURRENT ROW

        ) AS rolling_30d_sales

    FROM base

),
/* =====================================================
DAILY MOH
===================================================== */

daily_moh AS (

    SELECT

        store_code,

        dt,

        week_id,

        month_id,

        ROUND(

            daily_stock::numeric

            /

            CASE

                WHEN rolling_30d_sales <= 0

                THEN 1

                ELSE rolling_30d_sales

            END,

            2

        ) AS moh_daily

    FROM calc

)

/* =====================================================
FINAL OUTPUT
===================================================== */

SELECT

    d.store_code,

    sn.store_name,

    MAX(CASE WHEN d.dt = db.d1 THEN d.moh_daily END) AS d1_moh,
    MAX(CASE WHEN d.dt = db.d2 THEN d.moh_daily END) AS d2_moh,
    MAX(CASE WHEN d.dt = db.d3 THEN d.moh_daily END) AS d3_moh,
    MAX(CASE WHEN d.dt = db.d4 THEN d.moh_daily END) AS d4_moh,
    MAX(CASE WHEN d.dt = db.d5 THEN d.moh_daily END) AS d5_moh,
    MAX(CASE WHEN d.dt = db.d6 THEN d.moh_daily END) AS d6_moh,
    MAX(CASE WHEN d.dt = db.d7 THEN d.moh_daily END) AS d7_moh,

    MAX(CASE WHEN d.dt = db.w1 THEN d.moh_daily END) AS w1_moh,
    MAX(CASE WHEN d.dt = db.w2 THEN d.moh_daily END) AS w2_moh,
    MAX(CASE WHEN d.dt = db.w3 THEN d.moh_daily END) AS w3_moh,
    MAX(CASE WHEN d.dt = db.w4 THEN d.moh_daily END) AS w4_moh,

    MAX(CASE WHEN d.dt = db.m1 THEN d.moh_daily END) AS m1_moh,
    MAX(CASE WHEN d.dt = db.m2 THEN d.moh_daily END) AS m2_moh,
    MAX(CASE WHEN d.dt = db.m3 THEN d.moh_daily END) AS m3_moh,

    CURRENT_DATE AS update_date

FROM daily_moh d

CROSS JOIN date_bounds db

LEFT JOIN store_names sn
       ON d.store_code = sn.store_code

GROUP BY

    d.store_code,

    sn.store_name

ORDER BY

    d.store_code;  
    """
    },
        {
            "folder": "inventory_value",
            "file": "inventory_value_mrp_14_oct.xlsx",
            "query": """
    WITH inv AS (
    
        SELECT
            COALESCE(
                NULLIF(TRIM(scm.store_code::text), ''),
                TRIM(i.store_code::text)
            ) AS store_code,
    
            safe_to_date(LEFT(i.date::text, 10)) AS dt,
    
            /* NaN-safe value. If MRP value should be qty x price instead, use:
            GREATEST(COALESCE(i.total_stock,0),0)
            * CASE WHEN i.price::text IN ('NaN','Infinity','-Infinity') THEN 0
                    ELSE COALESCE(i.price,0)::numeric END               */
            SUM(
                CASE
                    WHEN i.stock_amt::text IN ('NaN','nan','Infinity','-Infinity')
                    THEN 0
                    ELSE COALESCE(i.stock_amt, 0)::numeric
                END
            ) AS inventory_value
    
        FROM inv_90_days i
    
        LEFT JOIN store_code_mapping scm
            ON TRIM(i.store_code::text) = TRIM(scm.old_store_code::text)
    
        WHERE COALESCE(i.total_stock, 0) >= 0
        AND COALESCE(i.category, '') NOT IN
            ('Consumable', 'Consumables', 'Decoration Materials')
        AND safe_to_date(LEFT(i.date::text, 10)) IS NOT NULL
    
        GROUP BY 1, 2
    ),
    
    date_bounds AS (
        SELECT
            (CURRENT_DATE - 1)::date AS d1,
            (CURRENT_DATE - 2)::date AS d2,
            (CURRENT_DATE - 3)::date AS d3,
            (CURRENT_DATE - 4)::date AS d4,
            (CURRENT_DATE - 5)::date AS d5,
            (CURRENT_DATE - 6)::date AS d6,
            (CURRENT_DATE - 7)::date AS d7,
            TO_CHAR(CURRENT_DATE - INTERVAL '1 week', 'IYYYIW')::bigint AS w1,
            TO_CHAR(CURRENT_DATE - INTERVAL '2 week', 'IYYYIW')::bigint AS w2,
            TO_CHAR(CURRENT_DATE - INTERVAL '3 week', 'IYYYIW')::bigint AS w3,
            TO_CHAR(CURRENT_DATE - INTERVAL '4 week', 'IYYYIW')::bigint AS w4,
            TO_CHAR(DATE_TRUNC('month', CURRENT_DATE) - INTERVAL '1 month', 'YYYYMM')::bigint AS m1,
            TO_CHAR(DATE_TRUNC('month', CURRENT_DATE) - INTERVAL '2 month', 'YYYYMM')::bigint AS m2,
            TO_CHAR(DATE_TRUNC('month', CURRENT_DATE) - INTERVAL '3 month', 'YYYYMM')::bigint AS m3
    ),
    
    /* ---------- DAILY (one row per store) ---------- */
    daily_p AS (
        SELECT
            iv.store_code,
            MAX(CASE WHEN iv.dt = db.d1 THEN iv.inventory_value END) AS day_1,
            MAX(CASE WHEN iv.dt = db.d2 THEN iv.inventory_value END) AS day_2,
            MAX(CASE WHEN iv.dt = db.d3 THEN iv.inventory_value END) AS day_3,
            MAX(CASE WHEN iv.dt = db.d4 THEN iv.inventory_value END) AS day_4,
            MAX(CASE WHEN iv.dt = db.d5 THEN iv.inventory_value END) AS day_5,
            MAX(CASE WHEN iv.dt = db.d6 THEN iv.inventory_value END) AS day_6,
            MAX(CASE WHEN iv.dt = db.d7 THEN iv.inventory_value END) AS day_7
        FROM inv iv
        CROSS JOIN date_bounds db
        GROUP BY iv.store_code
    ),
    
    /* ---------- WEEKLY AVERAGE (one row per store) ---------- */
    weekly_avg AS (
        SELECT
            store_code,
            TO_CHAR(dt, 'IYYYIW')::bigint AS week_id,
            AVG(inventory_value) AS inventory_value
        FROM inv
        GROUP BY 1, 2
    ),
    
    weekly_p AS (
        SELECT
            wa.store_code,
            MAX(CASE WHEN wa.week_id = db.w1 THEN wa.inventory_value END) AS week_1,
            MAX(CASE WHEN wa.week_id = db.w2 THEN wa.inventory_value END) AS week_2,
            MAX(CASE WHEN wa.week_id = db.w3 THEN wa.inventory_value END) AS week_3,
            MAX(CASE WHEN wa.week_id = db.w4 THEN wa.inventory_value END) AS week_4
        FROM weekly_avg wa
        CROSS JOIN date_bounds db
        GROUP BY wa.store_code
    ),
    
    /* ---------- MONTHLY AVERAGE (one row per store) ---------- */
    monthly_avg AS (
        SELECT
            store_code,
            TO_CHAR(dt, 'YYYYMM')::bigint AS month_id,
            AVG(inventory_value) AS inventory_value
        FROM inv
        GROUP BY 1, 2
    ),
    
    monthly_p AS (
        SELECT
            ma.store_code,
            MAX(CASE WHEN ma.month_id = db.m1 THEN ma.inventory_value END) AS month_1,
            MAX(CASE WHEN ma.month_id = db.m2 THEN ma.inventory_value END) AS month_2,
            MAX(CASE WHEN ma.month_id = db.m3 THEN ma.inventory_value END) AS month_3
        FROM monthly_avg ma
        CROSS JOIN date_bounds db
        GROUP BY ma.store_code
    )
    
    /* ---------- FINAL: 1 row per store, no row explosion ---------- */
    SELECT
        d.store_code,
        ROUND(d.day_1, 0)::BIGINT AS day_1_inventory_value,
        ROUND(d.day_2, 0)::BIGINT AS day_2_inventory_value,
        ROUND(d.day_3, 0)::BIGINT AS day_3_inventory_value,
        ROUND(d.day_4, 0)::BIGINT AS day_4_inventory_value,
        ROUND(d.day_5, 0)::BIGINT AS day_5_inventory_value,
        ROUND(d.day_6, 0)::BIGINT AS day_6_inventory_value,
        ROUND(d.day_7, 0)::BIGINT AS day_7_inventory_value,
        ROUND(w.week_1, 0)::BIGINT AS week_1_inventory_value,
        ROUND(w.week_2, 0)::BIGINT AS week_2_inventory_value,
        ROUND(w.week_3, 0)::BIGINT AS week_3_inventory_value,
        ROUND(w.week_4, 0)::BIGINT AS week_4_inventory_value,
        ROUND(m.month_1, 0)::BIGINT AS month_1_inventory_value,
        ROUND(m.month_2, 0)::BIGINT AS month_2_inventory_value,
        ROUND(m.month_3, 0)::BIGINT AS month_3_inventory_value,
        CURRENT_DATE AS update_date
    FROM daily_p d
    LEFT JOIN weekly_p  w ON d.store_code = w.store_code
    LEFT JOIN monthly_p m ON d.store_code = m.store_code
    ORDER BY d.store_code;
            """
        },
      {
        "folder": "inventory_value",
        "file": "inventory_value_cp_14_oct.xlsx",
        "query": """
           WITH price_master AS (

    SELECT

        TRIM(product_code::text) AS product_code,

        CASE

            WHEN b2b_avg_price IS NULL THEN 0

            WHEN b2b_avg_price::text IN
            (
                'NaN',
                'nan',
                'Infinity',
                '-Infinity'
            )
            THEN 0

            ELSE b2b_avg_price::numeric

        END AS b2b_avg_price,

        category

    FROM product_master

),

/* =====================================================
INVENTORY VALUE DAILY
===================================================== */

inventory_value_daily AS (

    SELECT

        COALESCE(

            NULLIF(TRIM(scm.store_code::text), ''),

            TRIM(i.store_code::text)

        ) AS store_code,

        CAST(i.date AS DATE) AS dt,

        SUM(

            COALESCE(i.total_stock,0)

            *

            COALESCE(pm.b2b_avg_price,0)

        ) AS inventory_value

    FROM inv_90_days i

    LEFT JOIN store_code_mapping scm

        ON TRIM(i.store_code::text)
         = TRIM(scm.old_store_code::text)

    INNER JOIN price_master pm

        ON TRIM(i.product_code::text)
         = pm.product_code

    WHERE

        /* Ignore only negative inventory */

        COALESCE(i.total_stock,0) >= 0

        /* Remove unwanted categories */

        AND COALESCE(pm.category,'') NOT IN
        (
            'Consumable',
            'Consumables',
            'Decoration Materials'
        )

    GROUP BY

        1,
        2

),

/* =====================================================
DATE BOUNDS
===================================================== */

date_bounds AS (

    SELECT

        CURRENT_DATE - INTERVAL '1 day' AS d1,

        CURRENT_DATE - INTERVAL '2 day' AS d2,

        CURRENT_DATE - INTERVAL '3 day' AS d3,

        CURRENT_DATE - INTERVAL '4 day' AS d4,

        CURRENT_DATE - INTERVAL '5 day' AS d5,

        CURRENT_DATE - INTERVAL '6 day' AS d6,

        CURRENT_DATE - INTERVAL '7 day' AS d7,

        /* WEEKLY SNAPSHOT DATES */

(
    CURRENT_DATE
    - EXTRACT(DOW FROM CURRENT_DATE)::int
) AS w1,

(
    CURRENT_DATE
    - EXTRACT(DOW FROM CURRENT_DATE)::int
    - INTERVAL '7 day'
) AS w2,

(
    CURRENT_DATE
    - EXTRACT(DOW FROM CURRENT_DATE)::int
    - INTERVAL '14 day'
) AS w3,

(
    CURRENT_DATE
    - EXTRACT(DOW FROM CURRENT_DATE)::int
    - INTERVAL '21 day'
) AS w4,

        /* MONTHLY SNAPSHOT DATES */

(
    DATE_TRUNC('month', CURRENT_DATE)
    - INTERVAL '1 day'
)::date AS m1,

(
    DATE_TRUNC('month', CURRENT_DATE)
    - INTERVAL '1 month'
    - INTERVAL '1 day'
)::date AS m2,

(
    DATE_TRUNC('month', CURRENT_DATE)
    - INTERVAL '2 month'
    - INTERVAL '1 day'
)::date AS m3

)

/* =====================================================
FINAL OUTPUT
===================================================== */

SELECT

    iv.store_code,

    ROUND(

        MAX(

            CASE

                WHEN iv.dt = db.d1

                THEN iv.inventory_value

            END

        ),

        0

    )::BIGINT AS day_1_inventory_value,

    ROUND(

        MAX(

            CASE

                WHEN iv.dt = db.d2

                THEN iv.inventory_value

            END

        ),

        0

    )::BIGINT AS day_2_inventory_value,

    ROUND(

        MAX(

            CASE

                WHEN iv.dt = db.d3

                THEN iv.inventory_value

            END

        ),

        0

    )::BIGINT AS day_3_inventory_value,

    ROUND(

        MAX(

            CASE

                WHEN iv.dt = db.d4

                THEN iv.inventory_value

            END

        ),

        0

    )::BIGINT AS day_4_inventory_value,

    ROUND(

        MAX(

            CASE

                WHEN iv.dt = db.d5

                THEN iv.inventory_value

            END

        ),

        0

    )::BIGINT AS day_5_inventory_value,

    ROUND(

        MAX(

            CASE

                WHEN iv.dt = db.d6

                THEN iv.inventory_value

            END

        ),

        0

    )::BIGINT AS day_6_inventory_value,

    ROUND(

        MAX(

            CASE

                WHEN iv.dt = db.d7

                THEN iv.inventory_value

            END

        ),

        0

    )::BIGINT AS day_7_inventory_value,
	    ROUND(
    MAX(
        CASE
            WHEN iv.dt = db.w1
            THEN iv.inventory_value
        END
    ),
    0
)::BIGINT AS week_1_inventory_value,

ROUND(
    MAX(
        CASE
            WHEN iv.dt = db.w2
            THEN iv.inventory_value
        END
    ),
    0
)::BIGINT AS week_2_inventory_value,

ROUND(
    MAX(
        CASE
            WHEN iv.dt = db.w3
            THEN iv.inventory_value
        END
    ),
    0
)::BIGINT AS week_3_inventory_value,

ROUND(
    MAX(
        CASE
            WHEN iv.dt = db.w4
            THEN iv.inventory_value
        END
    ),
    0
)::BIGINT AS week_4_inventory_value,

   ROUND(
    MAX(
        CASE
            WHEN iv.dt = db.m1
            THEN iv.inventory_value
        END
    ),
    0
)::BIGINT AS month_1_inventory_value,

ROUND(
    MAX(
        CASE
            WHEN iv.dt = db.m2
            THEN iv.inventory_value
        END
    ),
    0
)::BIGINT AS month_2_inventory_value,

ROUND(
    MAX(
        CASE
            WHEN iv.dt = db.m3
            THEN iv.inventory_value
        END
    ),
    0
)::BIGINT AS month_3_inventory_value,

    CURRENT_DATE AS update_date

FROM inventory_value_daily iv

CROSS JOIN date_bounds db

GROUP BY

    iv.store_code

ORDER BY

    iv.store_code;
        """
        },

      {
        "folder": "new_sku_availability",
        "file": "new_sku_availability_14_oct.xlsx",
        "query": """
        WITH calendar_dates AS (

    SELECT
        generate_series(
            CURRENT_DATE - INTERVAL '120 day',
            CURRENT_DATE - INTERVAL '1 day',
            INTERVAL '1 day'
        )::date AS report_date

),

/* =========================================================
STORE MAP
========================================================= */

store_map AS (

    SELECT
        TRIM(old_store_code::text) AS old_store_code,
        TRIM(store_code::text) AS new_store_code
    FROM store_code_mapping

),

/* =========================================================
MASTER SKU
========================================================= */

master_sku AS MATERIALIZED (

    SELECT DISTINCT

        TRIM(product_code::text) AS sku_code,

        safe_to_date(b2b_launch_date::text) AS launch_date

    FROM product_master

   WHERE b2b_launch_date IS NOT NULL
      AND safe_to_date(b2b_launch_date::text) IS NOT NULL

),

/* =========================================================
ORDERING MANIFEST
(Read only required dates)
========================================================= */

manifest_base AS MATERIALIZED (

    SELECT DISTINCT

        safe_to_date(date::text) AS manifest_date,

        TRIM(sku_code::text) AS sku_code

    FROM ordering_manifest

    WHERE date IS NOT NULL

      AND safe_to_date(date::text) IS NOT NULL

      AND safe_to_date(date::text) >= CURRENT_DATE - INTERVAL '124 day'

),

/* =========================================================
REPORT DATE × SKU PRESENT IN LAST 5 DAYS MOS

This replaces millions of EXISTS() calls later.
========================================================= */

manifest_last5 AS MATERIALIZED (

    SELECT DISTINCT

        cd.report_date,

        mb.sku_code

    FROM calendar_dates cd

    INNER JOIN manifest_base mb

        ON mb.manifest_date BETWEEN

           (cd.report_date - INTERVAL '5 day')::date
                   
           AND (cd.report_date - INTERVAL '1 day')::date
),

/* =========================================================
ELIGIBLE NEW SKU

Launch in last 30 days
AND present in last 5 day manifest
========================================================= */

eligible_new_sku_daily AS MATERIALIZED (

    SELECT DISTINCT

        cd.report_date,

        ms.sku_code

    FROM calendar_dates cd

    INNER JOIN master_sku ms

        ON ms.launch_date BETWEEN
           (cd.report_date - INTERVAL '30 day')::date
           AND cd.report_date

),
/* =========================================================
STORE STOCK
========================================================= */

stock_base AS MATERIALIZED (

    SELECT

        COALESCE(
            NULLIF(TRIM(sm.new_store_code), ''),
            NULLIF(TRIM(i.store_code::text), '')
        ) AS store_code,

       safe_to_date(i.date::text) AS stock_date,

        TRIM(i.product_code::text) AS sku_code,

        SUM(
            GREATEST(
                COALESCE(i.available_stock, 0),
                0
            )
        ) AS available_stock

    FROM inv_90_days i

    LEFT JOIN store_map sm

        ON TRIM(i.store_code::text) =
           TRIM(sm.old_store_code)

    WHERE COALESCE(
              NULLIF(TRIM(sm.new_store_code), ''),
              NULLIF(TRIM(i.store_code::text), '')
          ) IS NOT NULL

      AND safe_to_date(i.date::text) IS NOT NULL

      AND safe_to_date(i.date::text) >= CURRENT_DATE - INTERVAL '120 day'

    GROUP BY
        1,
        2,
        3

),

/* =========================================================
STORE LIST
========================================================= */

store_list AS (

    SELECT DISTINCT
        store_code
    FROM stock_base

),

/* =========================================================
STORE × ELIGIBLE SKU BASE

Every eligible SKU for every store
========================================================= */

store_sku_base AS (

    SELECT

        sl.store_code,

        ens.report_date,

        ens.sku_code

    FROM store_list sl

    CROSS JOIN eligible_new_sku_daily ens

),

/* =========================================================
SKU STATUS

Missing inventory = 0
No correlated EXISTS()
========================================================= */

sku_status AS (

    SELECT

        ssb.store_code,

        ssb.report_date,

        ssb.sku_code,

        COALESCE(
            sb.available_stock,
            0
        ) AS available_stock,

        CASE
            WHEN ml.sku_code IS NOT NULL
            THEN 1
            ELSE 0
        END AS in_mos

    FROM store_sku_base ssb

    LEFT JOIN stock_base sb

           ON sb.store_code = ssb.store_code
          AND sb.stock_date = ssb.report_date
          AND sb.sku_code = ssb.sku_code

    LEFT JOIN manifest_last5 ml

           ON ml.report_date = ssb.report_date
          AND ml.sku_code = ssb.sku_code

),
/* =========================================================
DAILY STORE AVAILABILITY
========================================================= */

daily_store_availability AS (

    SELECT

        store_code,

        report_date,

        COUNT(
            CASE
                WHEN available_stock > 2
                THEN 1
            END
        ) AS numerator_value,

        COUNT(*)
        -
        COUNT(
            CASE
                WHEN available_stock < 3
                 AND in_mos = 0
                THEN 1
            END
        ) AS denominator_value

    FROM sku_status

    GROUP BY
        store_code,
        report_date

),

/* =========================================================
DAILY AVAILABILITY
========================================================= */

daily_availability AS (

    SELECT

        store_code,

        report_date,

        numerator_value,

        denominator_value,

        ROUND(
            numerator_value::numeric
            /
            NULLIF(denominator_value, 0),
            4
        ) AS availability

    FROM daily_store_availability

),

/* =========================================================
DAILY PIVOT
========================================================= */

daily_pivot AS (

    SELECT

        store_code,

        ROUND(MAX(CASE WHEN report_date = CURRENT_DATE - INTERVAL '1 day' THEN availability END),4) AS d1,
        ROUND(MAX(CASE WHEN report_date = CURRENT_DATE - INTERVAL '2 day' THEN availability END),4) AS d2,
        ROUND(MAX(CASE WHEN report_date = CURRENT_DATE - INTERVAL '3 day' THEN availability END),4) AS d3,
        ROUND(MAX(CASE WHEN report_date = CURRENT_DATE - INTERVAL '4 day' THEN availability END),4) AS d4,
        ROUND(MAX(CASE WHEN report_date = CURRENT_DATE - INTERVAL '5 day' THEN availability END),4) AS d5,
        ROUND(MAX(CASE WHEN report_date = CURRENT_DATE - INTERVAL '6 day' THEN availability END),4) AS d6,
        ROUND(MAX(CASE WHEN report_date = CURRENT_DATE - INTERVAL '7 day' THEN availability END),4) AS d7

    FROM daily_availability

    GROUP BY
        store_code

),

/* =========================================================
WEEKLY AVERAGE
========================================================= */

weekly_avg AS (

    SELECT

        store_code,

        TO_CHAR(report_date,'IYYYIW') AS week_id,

        ROUND(
            AVG(availability),
            4
        ) AS availability

    FROM daily_availability

    GROUP BY
        store_code,
        TO_CHAR(report_date,'IYYYIW')

),

/* =========================================================
MONTHLY AVERAGE
========================================================= */

monthly_avg AS (

    SELECT

        store_code,

        TO_CHAR(report_date,'YYYYMM') AS month_id,

        ROUND(
            AVG(availability),
            4
        ) AS availability

    FROM daily_availability

    GROUP BY
        store_code,
        TO_CHAR(report_date,'YYYYMM')

)
/* =========================================================
FINAL OUTPUT
========================================================= */

SELECT

    dp.store_code,

    dp.d1,
    dp.d2,
    dp.d3,
    dp.d4,
    dp.d5,
    dp.d6,
    dp.d7,

    ROUND(
        MAX(
            CASE
                WHEN wa.week_id =
                     TO_CHAR(
                         CURRENT_DATE - INTERVAL '1 week',
                         'IYYYIW'
                     )
                THEN wa.availability
            END
        ),
        4
    ) AS w1,

    ROUND(
        MAX(
            CASE
                WHEN wa.week_id =
                     TO_CHAR(
                         CURRENT_DATE - INTERVAL '2 week',
                         'IYYYIW'
                     )
                THEN wa.availability
            END
        ),
        4
    ) AS w2,

    ROUND(
        MAX(
            CASE
                WHEN wa.week_id =
                     TO_CHAR(
                         CURRENT_DATE - INTERVAL '3 week',
                         'IYYYIW'
                     )
                THEN wa.availability
            END
        ),
        4
    ) AS w3,

    ROUND(
        MAX(
            CASE
                WHEN wa.week_id =
                     TO_CHAR(
                         CURRENT_DATE - INTERVAL '4 week',
                         'IYYYIW'
                     )
                THEN wa.availability
            END
        ),
        4
    ) AS w4,

    ROUND(
        MAX(
            CASE
                WHEN ma.month_id =
                     TO_CHAR(
                         DATE_TRUNC('month', CURRENT_DATE) - INTERVAL '1 month',
                         'YYYYMM'
                     )
                THEN ma.availability
            END
        ),
        4
    ) AS m1,

    ROUND(
        MAX(
            CASE
                WHEN ma.month_id =
                     TO_CHAR(
                         DATE_TRUNC('month', CURRENT_DATE) - INTERVAL '2 month',
                         'YYYYMM'
                     )
                THEN ma.availability
            END
        ),
        4
    ) AS m2,

    ROUND(
        MAX(
            CASE
                WHEN ma.month_id =
                     TO_CHAR(
                         DATE_TRUNC('month', CURRENT_DATE) - INTERVAL '3 month',
                         'YYYYMM'
                     )
                THEN ma.availability
            END
        ),
        4
    ) AS m3,

    CURRENT_DATE AS update_date

FROM daily_pivot dp

LEFT JOIN weekly_avg wa
       ON dp.store_code = wa.store_code

LEFT JOIN monthly_avg ma
       ON dp.store_code = ma.store_code

GROUP BY

    dp.store_code,

    dp.d1,
    dp.d2,
    dp.d3,
    dp.d4,
    dp.d5,
    dp.d6,
    dp.d7

ORDER BY

    dp.store_code;
        """
      },
      {
        "folder": "b_class",
        "file": "b_class_14_oct.xlsx",
        "query": """
        WITH date_limits AS (
        
            SELECT
                CURRENT_DATE - INTERVAL '180 day' AS sales_start_date,
                CURRENT_DATE - INTERVAL '150 day' AS stock_start_date
        
        ),
        
        /* =========================================================
        STORE MAP
        ========================================================= */
        
        store_map AS (
        
            SELECT
                old_store_code,
                store_code AS new_store_code
            FROM store_code_mapping
        
        ),
        
        /* =========================================================
        STORE NAMES
        ========================================================= */
        
        store_names AS (
        
            SELECT
        
                COALESCE(
                    sm.new_store_code,
                    si.store_code
                ) AS store_code,
        
                MAX(si.store_name) AS store_name
        
            FROM inv_90_days si
        
            LEFT JOIN store_map sm
        
                ON si.store_code = sm.old_store_code
        
            GROUP BY 1
        
        ),
        
        /* =========================================================
        ALL STORES
        ========================================================= */
        
        stores AS (
        
            SELECT DISTINCT
        
                COALESCE(
                    sm.new_store_code,
                    si.store_code
                ) AS store_code
        
            FROM inv_90_days si
        
            LEFT JOIN store_map sm
        
                ON si.store_code = sm.old_store_code
        
        ),
        
        /* =========================================================
        REPORTING DATES
        ========================================================= */
        
        reporting_dates AS (
        
            SELECT DISTINCT
                safe_to_date(LEFT(date,10)) AS report_date
            FROM inv_90_days
            WHERE safe_to_date(LEFT(date,10)) IS NOT NULL
              AND safe_to_date(LEFT(date,10)) >=
                  (SELECT stock_start_date::date
                   FROM date_limits)
        
        ),
        
        /* =========================================================
        DAILY SALES
        ========================================================= */
        
        daily_sales AS (
        
            SELECT
        
               safe_to_date(LEFT(date,10)) AS sales_date,
        
                TRIM(product_code) AS product_code,
        
                SUM(COALESCE(net_sales_qty,0)) AS sales_qty,
        
                SUM(COALESCE(net_sales,0)) AS sales_amount
        
            FROM last_3_month_sale
        
            WHERE safe_to_date(LEFT(date,10)) IS NOT NULL
              AND safe_to_date(LEFT(date,10)) >=
                  (SELECT sales_start_date::date
                   FROM date_limits)
        
              AND COALESCE(category,'')
        
                  NOT IN
        
                  (
        
                      'Consumable',
        
                      'Decoration Materials'
        
                  )
        
            GROUP BY
        
                1,
        
                2
        
        ),
        
        /* =========================================================
        ROLLING 60 DAY SALES
        ========================================================= */
        
        rolling_60d_sales AS (
        
            SELECT
        
                rd.report_date,
        
                ds.product_code,
        
                SUM(ds.sales_qty) AS sales_qty_60d,
        
                SUM(ds.sales_amount) AS sales_amount_60d
        
            FROM reporting_dates rd
        
            INNER JOIN daily_sales ds
        
                ON ds.sales_date BETWEEN
        
                   rd.report_date - INTERVAL '59 day'
        
                   AND rd.report_date
        
            GROUP BY
        
                1,
        
                2
        
        ),
        
        /* =========================================================
        QTY RANK
        ========================================================= */
        
        qty_ranked AS (
        
            SELECT
        
                *,
        
                ROW_NUMBER() OVER(
        
                    PARTITION BY report_date
        
                    ORDER BY
        
                        sales_qty_60d DESC,
        
                        product_code
        
                ) AS qty_rank
        
            FROM rolling_60d_sales
        
        ),
        
        /* =========================================================
        AMOUNT RANK
        ========================================================= */
        
        amount_ranked AS (
        
            SELECT
        
                *,
        
                ROW_NUMBER() OVER(
        
                    PARTITION BY report_date
        
                    ORDER BY
        
                        sales_amount_60d DESC,
        
                        product_code
        
                ) AS amount_rank
        
            FROM rolling_60d_sales
        
        ),
        
        /* =========================================================
        WEIGHTED RANK
        ========================================================= */
        
        weighted_ranked AS (
        
            SELECT
        
                q.report_date,
        
                q.product_code,
        
                ROUND(
        
                    (
        
                        q.qty_rank
        
                        +
        
                        a.amount_rank
        
                    ) / 2.0,
        
                    2
        
                ) AS weighted_rank
        
            FROM qty_ranked q
        
            INNER JOIN amount_ranked a
        
                ON q.report_date = a.report_date
        
               AND q.product_code = a.product_code
        
        ),
        
        /* =========================================================
        FINAL RANK
        ========================================================= */
        
        final_ranked_skus AS (
        
            SELECT
        
                report_date,
        
                product_code,
        
                ROW_NUMBER() OVER(
        
                    PARTITION BY report_date
        
                    ORDER BY
        
                        weighted_rank,
        
                        product_code
        
                ) AS final_rank
        
            FROM weighted_ranked
        
        ),
        
        /* =========================================================
        TARGET SKUS
        ========================================================= */
        
        target_skus AS (
        
            SELECT
        
                report_date,
        
                product_code
        
            FROM final_ranked_skus
        
            WHERE final_rank BETWEEN 301 AND 600
        
        ),
        /* =========================================================
        MANIFEST
        ========================================================= */
        
        manifest_base AS (
        
            SELECT DISTINCT
                safe_to_date(LEFT(date,10)) AS manifest_date,
                TRIM(sku_code) AS product_code
            FROM ordering_manifest
            WHERE safe_to_date(LEFT(date,10)) IS NOT NULL
        
        ),
        
        /* =========================================================
        ELIGIBLE TARGET SKUS
        ========================================================= */
        
        eligible_target_skus AS (
        
            SELECT DISTINCT
        
                ts.report_date,
        
                ts.product_code
        
            FROM target_skus ts
        
            INNER JOIN manifest_base mb
        
                ON ts.product_code = mb.product_code
        
               AND mb.manifest_date BETWEEN
        
                   ts.report_date - INTERVAL '5 day'
        
                   AND ts.report_date - INTERVAL '1 day'
        
        ),
        
        /* =========================================================
        STOCK FILTERED
        ========================================================= */
        
        stock_filtered AS (
        
            SELECT
        
                COALESCE(
                    sm.new_store_code,
                    si.store_code
                ) AS store_code,
        
                safe_to_date(LEFT(si.date,10)) AS stock_date,
        
                TRIM(si.product_code) AS product_code,
        
                SUM(
                    GREATEST(
                        COALESCE(si.available_stock,0),
                        0
                    )
                ) AS available_stock
        
            FROM inv_90_days si
        
            LEFT JOIN store_map sm
        
                ON si.store_code = sm.old_store_code
        
            WHERE safe_to_date(LEFT(si.date,10)) IS NOT NULL
              AND safe_to_date(LEFT(si.date,10)) >=
                  (
                      SELECT stock_start_date::date
                      FROM date_limits
                  )
        
              AND COALESCE(si.category,'')
                  NOT IN
                  (
                      'Consumable',
                      'Decoration Materials'
                  )
        
            GROUP BY
        
                1,
                2,
                3
        
        ),
        
        /* =========================================================
        TARGET STOCK
        ========================================================= */
        
        target_stock AS (
        
            SELECT
        
                st.store_code,
        
                ts.report_date AS stock_date,
        
                ts.product_code,
        
                COALESCE(sf.available_stock,0) AS available_stock,
        
                CASE
        
                    WHEN COALESCE(sf.available_stock,0) >= 3
        
                    THEN 1
        
                    ELSE 0
        
                END AS is_available,
        
                CASE
        
                    WHEN ets.product_code IS NOT NULL
        
                    THEN 1
        
                    ELSE 0
        
                END AS in_mos
        
            FROM target_skus ts
        
            CROSS JOIN stores st
        
            LEFT JOIN stock_filtered sf
        
                   ON sf.store_code = st.store_code
        
                  AND sf.stock_date = ts.report_date
        
                  AND sf.product_code = ts.product_code
        
            LEFT JOIN eligible_target_skus ets
        
                   ON ets.report_date = ts.report_date
        
                  AND ets.product_code = ts.product_code
        
        ),
        
        /* =========================================================
        DAILY STORE AVAILABILITY
        ========================================================= */
        
        daily_store_availability AS (
        
            SELECT
        
                store_code,
        
                stock_date,
        
                COUNT(
        
                    DISTINCT CASE
        
                        WHEN available_stock >= 3
        
                        THEN product_code
        
                    END
        
                ) AS numerator_skus,
        
                COUNT(DISTINCT product_code)
        
                -
        
                COUNT(
        
                    DISTINCT CASE
        
                        WHEN in_mos = 0
        
                         AND available_stock < 3
        
                        THEN product_code
        
                    END
        
                ) AS denominator_skus
        
            FROM target_stock
        
            GROUP BY
        
                store_code,
        
                stock_date
        
        ),
        /* =========================================================
        DAILY KPI
        ========================================================= */
        
        daily_kpi AS (
        
            SELECT
        
                store_code,
        
                stock_date,
        
                ROUND(
        
                    numerator_skus::numeric
        
                    /
        
                    NULLIF(denominator_skus,0),
        
                    4
        
                ) AS availability,
        
                TO_CHAR(stock_date,'IYYYIW') AS week_id,
        
                TO_CHAR(stock_date,'YYYYMM') AS month_id
        
            FROM daily_store_availability
        
        ),
        
        /* =========================================================
        WEEKLY AVERAGE
        ========================================================= */

        weekly_avg AS (

        SELECT
            store_code,
            week_id,

            ROUND(
                AVG(availability),
                4
            ) AS availability

        FROM daily_kpi

        GROUP BY
            store_code,
            week_id

        ),

        /* =========================================================
        MONTHLY AVERAGE
        ========================================================= */

        monthly_avg AS (

        SELECT
            store_code,
            month_id,

            ROUND(
                AVG(availability),
                4
            ) AS availability

        FROM daily_kpi

        GROUP BY
            store_code,
            month_id


        
             )
        /* =========================================================
        FINAL OUTPUT
        ========================================================= */
        
        SELECT
        
            dk.store_code,
        
            sn.store_name,
        
            ROUND(MAX(CASE
                WHEN dk.stock_date = CURRENT_DATE - INTERVAL '1 day'
                THEN dk.availability
            END),4) AS d1,
        
            ROUND(MAX(CASE
                WHEN dk.stock_date = CURRENT_DATE - INTERVAL '2 day'
                THEN dk.availability
            END),4) AS d2,
        
            ROUND(MAX(CASE
                WHEN dk.stock_date = CURRENT_DATE - INTERVAL '3 day'
                THEN dk.availability
            END),4) AS d3,
        
            ROUND(MAX(CASE
                WHEN dk.stock_date = CURRENT_DATE - INTERVAL '4 day'
                THEN dk.availability
            END),4) AS d4,
        
            ROUND(MAX(CASE
                WHEN dk.stock_date = CURRENT_DATE - INTERVAL '5 day'
                THEN dk.availability
            END),4) AS d5,
        
            ROUND(MAX(CASE
                WHEN dk.stock_date = CURRENT_DATE - INTERVAL '6 day'
                THEN dk.availability
            END),4) AS d6,
        
            ROUND(MAX(CASE
                WHEN dk.stock_date = CURRENT_DATE - INTERVAL '7 day'
                THEN dk.availability
            END),4) AS d7,
        
            ROUND(
                MAX(
                    CASE
                        WHEN wa.week_id =
                             TO_CHAR(
                                 CURRENT_DATE - INTERVAL '1 week',
                                 'IYYYIW'
                             )
                        THEN wa.availability
                    END
                ),
                4
            ) AS w1,
        
            ROUND(
                MAX(
                    CASE
                        WHEN wa.week_id =
                             TO_CHAR(
                                 CURRENT_DATE - INTERVAL '2 week',
                                 'IYYYIW'
                             )
                        THEN wa.availability
                    END
                ),
                4
            ) AS w2,
        
            ROUND(
                MAX(
                    CASE
                        WHEN wa.week_id =
                             TO_CHAR(
                                 CURRENT_DATE - INTERVAL '3 week',
                                 'IYYYIW'
                             )
                        THEN wa.availability
                    END
                ),
                4
            ) AS w3,
        
            ROUND(
                MAX(
                    CASE
                        WHEN wa.week_id =
                             TO_CHAR(
                                 CURRENT_DATE - INTERVAL '4 week',
                                 'IYYYIW'
                             )
                        THEN wa.availability
                    END
                ),
                4
            ) AS w4,
        
            ROUND(
                MAX(
                    CASE
                        WHEN ma.month_id =
                             TO_CHAR(
                                 DATE_TRUNC('month', CURRENT_DATE)
                                 - INTERVAL '1 month',
                                 'YYYYMM'
                             )
                        THEN ma.availability
                    END
                ),
                4
            ) AS m1,
        
            ROUND(
                MAX(
                    CASE
                        WHEN ma.month_id =
                             TO_CHAR(
                                 DATE_TRUNC('month', CURRENT_DATE)
                                 - INTERVAL '2 month',
                                 'YYYYMM'
                             )
                        THEN ma.availability
                    END
                ),
                4
            ) AS m2,
        
            ROUND(
                MAX(
                    CASE
                        WHEN ma.month_id =
                             TO_CHAR(
                                 DATE_TRUNC('month', CURRENT_DATE)
                                 - INTERVAL '3 month',
                                 'YYYYMM'
                             )
                        THEN ma.availability
                    END
                ),
                4
            ) AS m3,
        
            CURRENT_DATE AS update_date
        
        FROM daily_kpi dk
        
        LEFT JOIN weekly_avg wa
               ON dk.store_code = wa.store_code
        
        LEFT JOIN monthly_avg ma
               ON dk.store_code = ma.store_code
        
        LEFT JOIN store_names sn
               ON dk.store_code = sn.store_code
        
        GROUP BY
        
            dk.store_code,
            sn.store_name
        
        ORDER BY
        
            dk.store_code;
        """
      },
      {
         "folder": "a_class",
        "file": "a_class_14_oct.xlsx",
        "query": """
      WITH date_limits AS (

    SELECT
        CURRENT_DATE - INTERVAL '180 day' AS sales_start_date,
        CURRENT_DATE - INTERVAL '150 day' AS stock_start_date

),

/* =========================================================
STORE MAP
========================================================= */

store_map AS (

    SELECT
        old_store_code,
        store_code AS new_store_code
    FROM store_code_mapping

),

/* =========================================================
STORE NAMES
========================================================= */

store_names AS (

    SELECT

        COALESCE(
            sm.new_store_code,
            si.store_code
        ) AS store_code,

        MAX(si.store_name) AS store_name

    FROM inv_90_days si

    LEFT JOIN store_map sm

        ON si.store_code = sm.old_store_code

    GROUP BY 1

),

/* =========================================================
ALL STORES
========================================================= */

stores AS (

    SELECT DISTINCT

        COALESCE(
            sm.new_store_code,
            si.store_code
        ) AS store_code

    FROM inv_90_days si

    LEFT JOIN store_map sm

        ON si.store_code = sm.old_store_code

),

/* =========================================================
REPORTING DATES
========================================================= */

reporting_dates AS (

    SELECT DISTINCT

        safe_to_date(LEFT(date,10)) AS report_date

    FROM inv_90_days

    WHERE safe_to_date(LEFT(date,10)) IS NOT NULL

      AND safe_to_date(LEFT(date,10)) >=

          (SELECT stock_start_date::date
           FROM date_limits)

),

/* =========================================================
DAILY SALES
========================================================= */

daily_sales AS (

    SELECT

        safe_to_date(LEFT(date,10)) AS sales_date,

        TRIM(product_code) AS product_code,

        SUM(COALESCE(net_sales_qty,0)) AS sales_qty,

        SUM(COALESCE(net_sales,0)) AS sales_amount

    FROM last_3_month_sale

    WHERE safe_to_date(LEFT(date,10)) IS NOT NULL

      AND safe_to_date(LEFT(date,10)) >=

          (SELECT sales_start_date::date
           FROM date_limits)

      AND COALESCE(category,'')

          NOT IN

          (

              'Consumable',

              'Decoration Materials'

          )

    GROUP BY

        1,

        2

),

/* =========================================================
ROLLING 60 DAY SALES
========================================================= */

rolling_60d_sales AS (

    SELECT

        rd.report_date,

        ds.product_code,

        SUM(ds.sales_qty) AS sales_qty_60d,

        SUM(ds.sales_amount) AS sales_amount_60d

    FROM reporting_dates rd

    INNER JOIN daily_sales ds

        ON ds.sales_date BETWEEN

           rd.report_date - INTERVAL '59 day'

           AND rd.report_date

    GROUP BY

        1,

        2

),

/* =========================================================
QTY RANK
========================================================= */

qty_ranked AS (

    SELECT

        *,

        ROW_NUMBER() OVER(

            PARTITION BY report_date

            ORDER BY

                sales_qty_60d DESC,

                product_code

        ) AS qty_rank

    FROM rolling_60d_sales

),

/* =========================================================
AMOUNT RANK
========================================================= */

amount_ranked AS (

    SELECT

        *,

        ROW_NUMBER() OVER(

            PARTITION BY report_date

            ORDER BY

                sales_amount_60d DESC,

                product_code

        ) AS amount_rank

    FROM rolling_60d_sales

),

/* =========================================================
WEIGHTED RANK
========================================================= */

weighted_ranked AS (

    SELECT

        q.report_date,

        q.product_code,

        ROUND(

            (

                q.qty_rank

                +

                a.amount_rank

            ) / 2.0,

            2

        ) AS weighted_rank

    FROM qty_ranked q

    INNER JOIN amount_ranked a

        ON q.report_date = a.report_date

       AND q.product_code = a.product_code

),

/* =========================================================
FINAL RANK
========================================================= */

final_ranked_skus AS (

    SELECT

        report_date,

        product_code,

        ROW_NUMBER() OVER(

            PARTITION BY report_date

            ORDER BY

                weighted_rank,

                product_code

        ) AS final_rank

    FROM weighted_ranked

),

/* =========================================================
TARGET SKUS
========================================================= */

target_skus AS (

    SELECT

        report_date,

        product_code

    FROM final_ranked_skus

    WHERE final_rank <= 300

),
/* =========================================================
MANIFEST
========================================================= */

manifest_base AS (

    SELECT DISTINCT

        safe_to_date(LEFT(date,10)) AS manifest_date,

        TRIM(sku_code) AS product_code

    FROM ordering_manifest

    WHERE safe_to_date(LEFT(date,10)) IS NOT NULL

),

/* =========================================================
ELIGIBLE TARGET SKUS
========================================================= */

eligible_target_skus AS (

    SELECT DISTINCT

        ts.report_date,

        ts.product_code

    FROM target_skus ts

    INNER JOIN manifest_base mb

        ON ts.product_code = mb.product_code

       AND mb.manifest_date BETWEEN

           ts.report_date - INTERVAL '5 day'

           AND ts.report_date - INTERVAL '1 day'

),

/* =========================================================
STOCK FILTERED
========================================================= */

stock_filtered AS (

    SELECT

        COALESCE(
            sm.new_store_code,
            si.store_code
        ) AS store_code,

        safe_to_date(LEFT(si.date,10)) AS stock_date,

        TRIM(si.product_code) AS product_code,

        SUM(
            GREATEST(
                COALESCE(si.available_stock,0),
                0
            )
        ) AS available_stock

    FROM inv_90_days si

    LEFT JOIN store_map sm

        ON si.store_code = sm.old_store_code

    WHERE safe_to_date(LEFT(si.date,10)) IS NOT NULL

      AND safe_to_date(LEFT(si.date,10)) >=
          (
              SELECT stock_start_date::date
              FROM date_limits
          )

      AND COALESCE(si.category,'')
          NOT IN
          (
              'Consumable',
              'Decoration Materials'
          )

    GROUP BY

        1,
        2,
        3

),

/* =========================================================
TARGET STOCK
========================================================= */

target_stock AS (

    SELECT

        st.store_code,

        ts.report_date AS stock_date,

        ts.product_code,

        COALESCE(sf.available_stock,0) AS available_stock,

        CASE

            WHEN COALESCE(sf.available_stock,0) >= 3

            THEN 1

            ELSE 0

        END AS is_available,

        CASE

            WHEN ets.product_code IS NOT NULL

            THEN 1

            ELSE 0

        END AS in_mos

    FROM target_skus ts

    CROSS JOIN stores st

    LEFT JOIN stock_filtered sf

           ON sf.store_code = st.store_code

          AND sf.stock_date = ts.report_date

          AND sf.product_code = ts.product_code

    LEFT JOIN eligible_target_skus ets

           ON ets.report_date = ts.report_date

          AND ets.product_code = ts.product_code

),

/* =========================================================
DAILY STORE AVAILABILITY
========================================================= */

daily_store_availability AS (

    SELECT

        store_code,

        stock_date,

        COUNT(

            DISTINCT CASE

                WHEN available_stock >= 3

                THEN product_code

            END

        ) AS numerator_skus,

        COUNT(DISTINCT product_code)

        -

        COUNT(

            DISTINCT CASE

                WHEN in_mos = 0

                 AND available_stock < 3

                THEN product_code

            END

        ) AS denominator_skus

    FROM target_stock

    GROUP BY

        store_code,

        stock_date

),
/* =========================================================
DAILY KPI
========================================================= */

daily_kpi AS (

    SELECT

        store_code,

        stock_date,

        ROUND(

            numerator_skus::numeric

            /

            NULLIF(denominator_skus,0),

            4

        ) AS availability,

        TO_CHAR(stock_date,'IYYYIW') AS week_id,

        TO_CHAR(stock_date,'YYYYMM') AS month_id

    FROM daily_store_availability

),

/* =========================================================
WEEKLY AVERAGE
========================================================= */

weekly_avg AS (

    SELECT
        store_code,
        week_id,

        ROUND(
            AVG(availability),
            4
        ) AS availability

    FROM daily_kpi

    GROUP BY
        store_code,
        week_id

),

/* =========================================================
MONTHLY AVERAGE
========================================================= */

monthly_avg AS (

    SELECT
        store_code,
        month_id,

        ROUND(
            AVG(availability),
            4
        ) AS availability

    FROM daily_kpi

    GROUP BY
        store_code,
        month_id



)
/* =========================================================
FINAL OUTPUT
========================================================= */

SELECT

    dk.store_code,

    sn.store_name,

    ROUND(MAX(CASE
        WHEN dk.stock_date = CURRENT_DATE - INTERVAL '1 day'
        THEN dk.availability
    END),4) AS d1,

    ROUND(MAX(CASE
        WHEN dk.stock_date = CURRENT_DATE - INTERVAL '2 day'
        THEN dk.availability
    END),4) AS d2,

    ROUND(MAX(CASE
        WHEN dk.stock_date = CURRENT_DATE - INTERVAL '3 day'
        THEN dk.availability
    END),4) AS d3,

    ROUND(MAX(CASE
        WHEN dk.stock_date = CURRENT_DATE - INTERVAL '4 day'
        THEN dk.availability
    END),4) AS d4,

    ROUND(MAX(CASE
        WHEN dk.stock_date = CURRENT_DATE - INTERVAL '5 day'
        THEN dk.availability
    END),4) AS d5,

    ROUND(MAX(CASE
        WHEN dk.stock_date = CURRENT_DATE - INTERVAL '6 day'
        THEN dk.availability
    END),4) AS d6,

    ROUND(MAX(CASE
        WHEN dk.stock_date = CURRENT_DATE - INTERVAL '7 day'
        THEN dk.availability
    END),4) AS d7,

    ROUND(
        MAX(
            CASE
                WHEN wa.week_id =
                     TO_CHAR(
                         CURRENT_DATE - INTERVAL '1 week',
                         'IYYYIW'
                     )
                THEN wa.availability
            END
        ),
        4
    ) AS w1,

    ROUND(
        MAX(
            CASE
                WHEN wa.week_id =
                     TO_CHAR(
                         CURRENT_DATE - INTERVAL '2 week',
                         'IYYYIW'
                     )
                THEN wa.availability
            END
        ),
        4
    ) AS w2,

    ROUND(
        MAX(
            CASE
                WHEN wa.week_id =
                     TO_CHAR(
                         CURRENT_DATE - INTERVAL '3 week',
                         'IYYYIW'
                     )
                THEN wa.availability
            END
        ),
        4
    ) AS w3,

    ROUND(
        MAX(
            CASE
                WHEN wa.week_id =
                     TO_CHAR(
                         CURRENT_DATE - INTERVAL '4 week',
                         'IYYYIW'
                     )
                THEN wa.availability
            END
        ),
        4
    ) AS w4,

    ROUND(
        MAX(
            CASE
                WHEN ma.month_id =
                     TO_CHAR(
                         DATE_TRUNC('month', CURRENT_DATE)
                         - INTERVAL '1 month',
                         'YYYYMM'
                     )
                THEN ma.availability
            END
        ),
        4
    ) AS m1,

    ROUND(
        MAX(
            CASE
                WHEN ma.month_id =
                     TO_CHAR(
                         DATE_TRUNC('month', CURRENT_DATE)
                         - INTERVAL '2 month',
                         'YYYYMM'
                     )
                THEN ma.availability
            END
        ),
        4
    ) AS m2,

    ROUND(
        MAX(
            CASE
                WHEN ma.month_id =
                     TO_CHAR(
                         DATE_TRUNC('month', CURRENT_DATE)
                         - INTERVAL '3 month',
                         'YYYYMM'
                     )
                THEN ma.availability
            END
        ),
        4
    ) AS m3,

    CURRENT_DATE AS update_date

FROM daily_kpi dk

LEFT JOIN weekly_avg wa
       ON dk.store_code = wa.store_code

LEFT JOIN monthly_avg ma
       ON dk.store_code = ma.store_code

LEFT JOIN store_names sn
       ON dk.store_code = sn.store_code

GROUP BY

    dk.store_code,
    sn.store_name

ORDER BY

    dk.store_code;
    """
    },

]

# ==========================================
# Generate Reports
# ==========================================

collected_reports = {}   # filled by the loop, used for the dashboard JSON

for report in reports:

    try:

        print(f"\n▶ Running Report : {report['file']}")

        # Create folder if not exists
        folder_path = os.path.join(base_folder, report["folder"])
        os.makedirs(folder_path, exist_ok=True)

        # Execute SQL Query
        df = pd.read_sql_query(
            report["query"],
            engine
        )

        print(f"✅ Rows Retrieved : {len(df):,}")

        # Create timestamp file name
        file_name = report["file"].replace(
            ".xlsx",
            f"_{timestamp}.xlsx"
        )

        # Full output path
        file_path = os.path.join(
            folder_path,
            file_name
        )

        # Save Excel
        with pd.ExcelWriter(
            file_path,
            engine="openpyxl"
        ) as writer:

            df.to_excel(
                writer,
                sheet_name="Data",
                index=False
            )

        print(f"✅ Saved Successfully")
        print(file_path)

        collected_reports[report["file"]] = (df, file_path)

    except Exception:

        print(f"\n❌ ERROR IN REPORT : {report['file']}")
        traceback.print_exc()

# ==========================================
# Dashboard JSON (merch_kpi_data.json)
# ==========================================
# Written next to this script. Atomic: if anything fails, the previous
# JSON stays untouched.

JSON_OUTPUT_PATH = os.environ.get(
    "MERCH_KPI_JSON_PATH",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "merch_kpi_data.json"),
)

try:
    for note in merch_kpi_export.export(collected_reports, JSON_OUTPUT_PATH):
        print(f"   ⚠  {note}")
    print(f"\n✅ Dashboard JSON refreshed\n   {JSON_OUTPUT_PATH}")
except Exception:
    print("\n❌ Dashboard JSON NOT refreshed (previous JSON left untouched)")
    traceback.print_exc()

# ==========================================
# Close Connection
# ==========================================

engine.dispose()

print("\n🎉 All Reports Generated Successfully")