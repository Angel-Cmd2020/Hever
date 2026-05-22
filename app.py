import os
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# הגדרות העמוד (חובה להיות הפקודה הראשונה)
st.set_page_config(page_title="הטבות חבר | חיפוש", page_icon="💳", layout="wide")

# ==========================================
# עיצוב מותאם אישית - Dark Mode ועיצוב כרטיסיות למובייל
# ==========================================
st.markdown("""
<style>
    /* צביעת רקע האתר כולו לשחור עמוק */
    .stApp {
        background-color: #09090B !important;
    }
    
    /* מירכוז כללי והגבלת רוחב */
    .block-container {
        direction: rtl;
        max-width: 850px !important;
        margin: 0 auto !important;
        padding-top: 2rem !important;
    }
    
    /* עיצוב הכותרת הראשית */
    h1 {
        color: #8B5CF6 !important;
        text-align: center !important;
        margin-bottom: 30px !important;
    }
    
    /* עיצוב תווית החיפוש */
    .search-title {
        color: #C084FC !important;
        text-align: center !important;
        display: block !important;
        width: 100% !important;
        font-size: 18px !important;
        margin-bottom: 10px !important;
    }

    /* עיצוב תיבת הטקסט - מותאם למובייל וללא רשימות צפות */
    div[data-baseweb="input"] {
        background-color: #18181B !important;
        border: 2px solid #8B5CF6 !important;
        border-radius: 8px !important;
    }
    div[data-baseweb="input"] input {
        color: #A78BFA !important;
        font-size: 18px !important;
        font-weight: bold !important;
        text-align: center !important;
    }

    /* =========================================
       עיצוב טבלה - מצב תצוגה למחשב
       ========================================= */
    .custom-table {
        width: 100%;
        border-collapse: separate; /* שינוי שעוזר למנוע התנגשויות גבולות */
        border-spacing: 0;
        font-size: 16px;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        margin: 25px 0;
        border-radius: 8px 8px 0 0;
        overflow: hidden;
        box-shadow: 0 0 20px rgba(0, 0, 0, 0.5);
        border: none !important;
    }
    
    .custom-table th {
        background-color: #111827;
        color: #8B5CF6;
        text-align: center;
        font-weight: bold;
        padding: 15px;
        border-bottom: 2px solid #8B5CF6;
        border-left: none !important;
        border-right: none !important;
    }
    
    .custom-table td {
        padding: 14px 15px;
        text-align: center;
        color: #FFFFFF !important;
        font-weight: bold !important;
        border-bottom: 1px solid #27272A;
        border-left: none !important;  /* מחיקת הקו האנכי בצד שמאל */
        border-right: none !important; /* מחיקת הקו האנכי בצד ימין */
    }
    
    .custom-table tbody tr { background-color: #18181B; }
    .custom-table tbody tr:nth-of-type(odd) { background-color: #27272A; }
    .custom-table tbody tr:hover { background-color: #3B82F6; }
    .custom-table tbody tr:hover td { color: #FFFFFF !important; }

    /* =========================================
       התאמות מיוחדות למסכי פלאפון - עיצוב כרטיסיות
       ========================================= */
    @media (max-width: 600px) {
        .custom-table, .custom-table tbody, .custom-table tr, .custom-table td {
            display: block;
            width: 100%;
        }
        
        .custom-table thead {
            display: none; 
        }
        
        .custom-table tr {
            margin-bottom: 20px;
            border: 2px solid #8B5CF6 !important;
            border-radius: 12px;
            background-color: #18181B !important; 
            padding: 15px 10px;
            box-shadow: 0 4px 10px rgba(0,0,0,0.3);
        }
        
        .custom-table td {
            text-align: center !important;
            border-bottom: 1px solid #27272A !important;
            padding: 12px 5px !important;
            font-size: 18px !important;
            border-left: none !important;  /* וידוא מחיקת קווים צדדיים בתוך הכרטיסייה */
            border-right: none !important;
        }
        
        .custom-table td:last-child {
            border-bottom: none !important; 
        }
        
        .custom-table td::before {
            display: block;
            color: #C084FC;
            font-size: 14px;
            margin-bottom: 5px;
            font-weight: normal;
        }
        
        .custom-table td:nth-of-type(1)::before { content: "שם העסק / הרשת"; }
        .custom-table td:nth-of-type(2)::before { content: "הנחה: חבר שחור"; }
        .custom-table td:nth-of-type(3)::before { content: "הנחה: חבר ירוק"; }
    }
</style>
""", unsafe_allow_html=True)

# פונקציה לטעינת הנתונים
@st.cache_data
def load_and_process_data():
    EXCEL_PATH = "hever_discounts.xlsx"
    combined_data = {}
    
    if not os.path.exists(EXCEL_PATH):
        st.error(f"שגיאה: קובץ האקסל לא נמצא בנתיב: {EXCEL_PATH}")
        return pd.DataFrame()

    try:
        exclude_keywords = [
            "שם טכני", "tbl_", "רשתות / חנות", "רשתות / חנויות", 
            "מסעדות / אוכל / בילוי", "מסעדות / אוכל", "חבר ירוק", "חבר שחור",
            "שם חנות / מקום", "פאבים / ברי אוכל / בילוי", "קטגוריה",
            "שם רשת", "אטרקציות", "בילוי ופנאי", "ספא", "בתי קפה",
            "הטבות כרטיס שחור", "תינוקות וילדים", "אוכל", "אופנה", "בריאות ויופי",
            "שם חנות / רשת", "אחוז הנחה", "הנחה", "מדריך", "עמוד מס"
        ]

        df_black = pd.read_excel(EXCEL_PATH, sheet_name="חבר שחור")
        for _, row in df_black.iterrows():
            if len(row) >= 6:
                val_c = row.iloc[2]
                val_f = row.iloc[5]
                if pd.notna(val_c):
                    name = str(val_c).strip()
                    if len(name) > 1 and not name.isdigit() and not any(kw in name for kw in exclude_keywords):
                        discount_black = str(val_f).strip() if pd.notna(val_f) else "-"
                        if discount_black.lower() == "nan" or not discount_black:
                            discount_black = "-"
                        combined_data[name] = {"חבר ירוק": "-", "חבר שחור": discount_black}

        df_green = pd.read_excel(EXCEL_PATH, sheet_name="חבר ירוק")
        for _, row in df_green.iterrows():
            if len(row) >= 6:
                val_c = row.iloc[2]
                val_f = row.iloc[5]
                if pd.notna(val_c):
                    name = str(val_c).strip()
                    if len(name) > 1 and not name.isdigit() and not any(keyword in name for keyword in exclude_keywords):
                        discount_green = str(val_f).strip() if pd.notna(val_f) else "חבר ירוק"
                        if discount_green.lower() == "nan" or not discount_green:
                            discount_green = "חבר ירוק"
                        
                        if name in combined_data:
                            combined_data[name]["חבר ירוק"] = discount_green
                        else:
                            combined_data[name] = {"חבר ירוק": discount_green, "חבר שחור": "-"}

        if combined_data:
            df = pd.DataFrame.from_dict(combined_data, orient='index').reset_index()
            df.columns = ["שם העסק / הרשת", "הנחה: חבר ירוק", "הנחה: חבר שחור"]
            df = df[["שם העסק / הרשת", "הנחה: חבר שחור", "הנחה: חבר ירוק"]]
            return df
        return pd.DataFrame()

    except Exception as e:
        st.error(f"התרחשה שגיאה בעת קריאת הנתונים: {str(e)}")
        return pd.DataFrame()

# ==========================================
# בניית ממשק המשתמש (UI)
# ==========================================

st.markdown("<h1>Hever Discount Explorer 💳</h1>", unsafe_allow_html=True)

df_main = load_and_process_data()

if not df_main.empty:
    col1, col2, col3 = st.columns([1, 8, 1]) 
    with col2:
        st.markdown('<div class="search-title">הקש שם רשת / מסעדה / חנות לחיפוש:</div>', unsafe_allow_html=True)
        
        search_query = st.text_input(
            label="hidden_label", 
            placeholder="התחל להקליד...",
            label_visibility="collapsed"
        )

    if search_query:
        filtered_df = df_main[df_main["שם העסק / הרשת"].str.contains(search_query, case=False, na=False)]
    else:
        filtered_df = df_main

    html_table = filtered_df.to_html(classes='custom-table', index=False, escape=False)
    st.markdown(html_table, unsafe_allow_html=True)

    components.html(
        """
        <script>
        const parentDoc = window.parent.document;
        if (!parentDoc.getElementById("enter-key-closer")) {
            const script = parentDoc.createElement("script");
            script.id = "enter-key-closer";
            script.type = "text/javascript";
            script.innerHTML = `
                document.addEventListener('keydown', function(e) {
                    if (e.key === 'Enter' || e.keyCode === 13) {
                        setTimeout(function() {
                            let activeElem = document.activeElement;
                            if (activeElem && activeElem.tagName === 'INPUT') {
                                activeElem.blur();
                            }
                        }, 50);
                    }
                });
            `;
            parentDoc.head.appendChild(script);
        }
        </script>
        """,
        height=0
    )