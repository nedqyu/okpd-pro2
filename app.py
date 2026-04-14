"""
ОКПД 2 — Сопоставление кодов с ПП 1875 (Приложения №1 и №2)
Запуск: streamlit run app.py
"""

import re
import io
import streamlit as st
import pandas as pd

# ─────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────
st.set_page_config(
    page_title="ОКПД 2 · ПП 1875",
    page_icon="🗂️",
    layout="wide",
)

# ─────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600&family=Manrope:wght@400;500;600;700;800&display=swap');

html, body, [data-testid="stAppViewContainer"] {
    background: #0F172A !important;
    color: #CBD5E1;
    font-family: 'Manrope', sans-serif;
}
[data-testid="stSidebar"]       { background: #090E1A !important; }
[data-testid="stHeader"]        { background: transparent !important; }
h1,h2,h3                        { color: #F1F5F9 !important; font-family: 'Manrope', sans-serif; font-weight: 800; }

[data-baseweb="tab-list"]       { background: #1E293B !important; border-radius: 12px; padding: 4px; gap: 4px; }
[data-baseweb="tab"]            { border-radius: 8px !important; color: #94A3B8 !important; font-weight: 600 !important; }
[aria-selected="true"]          { background: #334155 !important; color: #F1F5F9 !important; }

input, textarea, [data-baseweb="input"] input {
    background: #1E293B !important; border: 1px solid #334155 !important;
    color: #F1F5F9 !important; border-radius: 8px !important;
    font-family: 'JetBrains Mono', monospace !important;
}

[data-testid="baseButton-primary"] {
    background: linear-gradient(135deg,#2563EB,#1D4ED8) !important;
    border: none !important; color: white !important; font-weight: 700 !important;
    border-radius: 8px !important; box-shadow: 0 4px 15px rgba(37,99,235,.35) !important;
    transition: all .2s !important;
}
[data-testid="baseButton-primary"]:hover {
    background: linear-gradient(135deg,#1D4ED8,#1E40AF) !important;
    transform: translateY(-1px); box-shadow: 0 6px 20px rgba(37,99,235,.5) !important;
}

[data-testid="baseButton-secondary"] {
    background: transparent !important; border: 1px solid #334155 !important;
    color: #64748B !important; border-radius: 8px !important; transition: all .2s !important;
}
[data-testid="baseButton-secondary"]:hover {
    background: rgba(239,68,68,.15) !important; border-color: #EF4444 !important; color: #EF4444 !important;
}

[data-testid="stDownloadButton"] button {
    background: linear-gradient(135deg,#059669,#047857) !important;
    border: none !important; color: white !important; font-weight: 700 !important;
    border-radius: 8px !important; box-shadow: 0 4px 15px rgba(5,150,105,.35) !important;
}
[data-testid="stDownloadButton"] button:hover {
    background: linear-gradient(135deg,#047857,#065F46) !important; transform: translateY(-1px);
}

[data-testid="stMetric"]      { background: #1E293B !important; border: 1px solid #334155 !important; border-radius: 12px !important; padding: 18px !important; }
[data-testid="stMetricValue"] { color: #38BDF8 !important; font-weight: 800 !important; }
[data-testid="stMetricLabel"] { color: #94A3B8 !important; }

[data-testid="stFileUploader"] {
    background: #1E293B !important; border: 2px dashed #334155 !important;
    border-radius: 12px !important; padding: 16px !important;
}

.card {
    background: #1E293B; border: 1px solid #334155; border-left: 4px solid #2563EB;
    border-radius: 12px; padding: 16px 20px; margin: 8px 0;
}
.card.yes  { border-left-color: #10B981; }
.card.no   { border-left-color: #EF4444; }
.card-label { color: #64748B; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing:.08em; margin-bottom:4px; }
.card-title { color: #F1F5F9; font-size: 18px; font-weight: 800; margin-bottom: 6px; }
.card-sub   { color: #94A3B8; font-size: 13px; font-family: 'JetBrains Mono', monospace; }
.card-badge { display: inline-block; padding: 2px 10px; border-radius: 20px; font-size: 12px; font-weight: 700; margin-left: 8px; }
.badge-yes  { background: rgba(16,185,129,.15); color: #10B981; border: 1px solid #10B981; }
.badge-no   { background: rgba(239,68,68,.12); color: #EF4444; border: 1px solid #EF4444; }
.mono { font-family: 'JetBrains Mono', monospace; color: #38BDF8; background: #0F172A; padding: 2px 8px; border-radius: 4px; font-size: 13px; }
hr   { border-color: #1E293B !important; }
::-webkit-scrollbar       { width:6px; height:6px; }
::-webkit-scrollbar-track { background: #0F172A; }
::-webkit-scrollbar-thumb { background: #334155; border-radius: 3px; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────
# УТИЛИТЫ
# ─────────────────────────────────────────────────────────

def clean_code(raw) -> str:
    return re.sub(r'[^\d.]', '', str(raw)).strip('.')


def load_main(file_obj) -> dict:
    """
    Главная база ОКПД 2.
    Фактическая структура: 'Код ОКПД 2' = порядковый №,
    'Статус' = сам код ОКПД, 'Название' = наименование.
    """
    df = pd.read_excel(file_obj, dtype=str)
    df.columns = [c.strip() for c in df.columns]

    # Определяем колонки гибко
    code_col = next((c for c in df.columns if c == 'Статус'), None)
    if code_col is None:
        code_col = next((c for c in df.columns if 'статус' in c.lower()), None)
    if code_col is None:
        # Если нет Статус — берём первую числово-кодовую колонку
        code_col = next((c for c in df.columns if 'код' in c.lower()), df.columns[1])

    name_col = next((c for c in df.columns if 'назван' in c.lower() or 'наимен' in c.lower()), df.columns[2])

    result = {}
    for _, row in df.iterrows():
        c = clean_code(row[code_col])
        n = str(row[name_col]).strip()
        if c and n.lower() not in ('nan', 'none', ''):
            result[c] = n
    return result


def load_appendix(file_obj) -> dict:
    """
    Приложение к ПП 1875.
    Строка 0 — заголовок, пропускаем (skiprows=1).
    col[0]=номер п/п, col[1]=наименование, col[2]=код ОКПД 2.
    """
    df = pd.read_excel(file_obj, header=None, dtype=str, skiprows=1)
    result = {}
    for _, row in df.iterrows():
        if len(row) < 3:
            continue
        c = clean_code(row.iloc[2])
        n = str(row.iloc[1]).strip()
        if c and n.lower() not in ('nan', 'none', ''):
            result[c] = n
    return result


def hier_lookup(user_code: str, lookup: dict) -> tuple:
    """
    1. Точное совпадение.
    2. user_code начинается на ключ словаря (пользователь длиннее — ищем корень).
    3. Ключ словаря начинается на user_code (пользователь короче — группа).
    """
    if not user_code:
        return False, ''
    if user_code in lookup:
        return True, lookup[user_code]
    best, best_len = '', 0
    for key in lookup:
        if user_code.startswith(key) and len(key) > best_len:
            best, best_len = key, len(key)
    if best:
        return True, lookup[best]
    for key in lookup:
        if key.startswith(user_code):
            return True, lookup[key]
    return False, ''


def excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        df.to_excel(w, index=False)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────
for key in ('hist_code', 'hist_name'):
    if key not in st.session_state:
        st.session_state[key] = []


# ─────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📁 Базы данных ПП 1875")

    f_main = st.file_uploader("📚 Справочник ОКПД 2 (.xlsx)", type=["xlsx","xls"], key="f_main")
    f_app1 = st.file_uploader("📋 Приложение №1 (.xlsx)",     type=["xlsx","xls"], key="f_app1")
    f_app2 = st.file_uploader("📋 Приложение №2 (.xlsx)",     type=["xlsx","xls"], key="f_app2")

    bases_ok = bool(f_main and f_app1 and f_app2)
    if bases_ok:
        st.success("✅ Все три базы загружены")
    else:
        missing = [n for f,n in [(f_main,"Справочник"),(f_app1,"Приложение №1"),(f_app2,"Приложение №2")] if not f]
        st.warning("Не загружено: " + ", ".join(missing))

    st.markdown("---")
    st.caption("Постановление Правительства РФ № 1875\nОКПД 2 — ОК 034-2014 (КПЕС 2008)")


# ─────────────────────────────────────────────────────────
# КЭШ БАЗ
# ─────────────────────────────────────────────────────────
@st.cache_data(show_spinner="⏳ Загружаю базы данных…")
def get_bases(m, a1, a2):
    return load_main(io.BytesIO(m)), load_appendix(io.BytesIO(a1)), load_appendix(io.BytesIO(a2))

main_d, app1_d, app2_d = {}, {}, {}
if bases_ok:
    main_d, app1_d, app2_d = get_bases(
        f_main.getvalue(), f_app1.getvalue(), f_app2.getvalue()
    )


# ─────────────────────────────────────────────────────────
# ЗАГОЛОВОК
# ─────────────────────────────────────────────────────────
st.markdown("""
<h1 style="margin-bottom:2px">🗂️ ОКПД 2 · Постановление № 1875</h1>
<p style="color:#64748B;margin-top:0;font-size:15px">
Сопоставление кодов со Справочником ОКПД 2 и Приложениями №1, №2
</p>
""", unsafe_allow_html=True)
st.markdown("---")


# ─────────────────────────────────────────────────────────
# ВКЛАДКИ
# ─────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📂 Массовая проверка", "🔎 Поиск по коду", "🔤 Поиск по названию"])


# ══════════════════════════════════════════════════════════
# TAB 1 — МАССОВАЯ ПРОВЕРКА
# ══════════════════════════════════════════════════════════
with tab1:
    st.markdown("### Загрузите рабочий файл")
    st.caption("Файл может содержать любое количество колонок. Обязательна колонка **Код ОКПД 2**.")

    user_file = st.file_uploader("Рабочий файл пользователя (.xlsx)", type=["xlsx","xls"], key="user_f")

    if user_file:
        if not bases_ok:
            st.warning("⚠️ Сначала загрузите все три базы в боковой панели.")
        else:
            user_df = pd.read_excel(user_file, dtype=str)
            user_df.columns = [str(c).strip() for c in user_df.columns]

            okpd_col = next(
                (c for c in user_df.columns if re.search(r'окпд.*2|код.*окпд', c.lower())),
                next((c for c in user_df.columns if 'окпд' in c.lower()), None)
            )

            if okpd_col is None:
                st.error(f"❌ Колонка 'Код ОКПД 2' не найдена. Найдены: {list(user_df.columns)}")
            else:
                col_l, col_r = st.columns([3,1])
                col_l.info(f"Колонка кодов: **{okpd_col}** · Строк: **{len(user_df)}**")

                if col_r.button("🚀 Сопоставить", type="primary"):
                    names_col, app1_col, app1_name_col, app2_col, app2_name_col = [], [], [], [], []
                    bar   = st.progress(0, text="Обрабатываю строки…")
                    total = len(user_df)

                    for i, raw in enumerate(user_df[okpd_col]):
                        code = clean_code(raw)
                        fm, vm = hier_lookup(code, main_d)
                        names_col.append(vm if fm else '')
                        f1, v1 = hier_lookup(code, app1_d)
                        app1_col.append('ДА' if f1 else 'НЕТ')
                        app1_name_col.append(v1 if f1 else '')
                        f2, v2 = hier_lookup(code, app2_d)
                        app2_col.append('ДА' if f2 else 'НЕТ')
                        app2_name_col.append(v2 if f2 else '')
                        if i % 50 == 0:
                            bar.progress(min(i/total, 1.0), text=f"Обработано {i}/{total}…")
                    bar.empty()

                    result = user_df.copy()
                    result['Наименование ОКПД 2']      = names_col
                    result['В Прил. №1 (ПП 1875)']     = app1_col
                    result['Наим. по Прил. №1']         = app1_name_col
                    result['В Прил. №2 (ПП 1875)']     = app2_col
                    result['Наим. по Прил. №2']         = app2_name_col

                    c1,c2,c3,c4,c5 = st.columns(5)
                    c1.metric("📋 Всего строк",       total)
                    c2.metric("📚 Найдено в ОКПД 2",  sum(1 for v in names_col if v),
                               f"{sum(1 for v in names_col if v)/total*100:.0f}%")
                    c3.metric("✅ В Приложении №1",   app1_col.count('ДА'),
                               f"{app1_col.count('ДА')/total*100:.0f}%")
                    c4.metric("✅ В Приложении №2",   app2_col.count('ДА'),
                               f"{app2_col.count('ДА')/total*100:.0f}%")
                    both = sum(1 for a,b in zip(app1_col,app2_col) if a=='ДА' and b=='ДА')
                    c5.metric("🔗 В обоих Прил.",     both, f"{both/total*100:.0f}%")

                    st.markdown("#### Предпросмотр (первые 5 строк)")
                    preview = list(user_df.columns) + [
                        'Наименование ОКПД 2','В Прил. №1 (ПП 1875)',
                        'Наим. по Прил. №1','В Прил. №2 (ПП 1875)','Наим. по Прил. №2'
                    ]
                    st.dataframe(result[preview].head(5), use_container_width=True)

                    st.download_button(
                        "⬇️ Скачать результат (.xlsx)",
                        data=excel_bytes(result),
                        file_name="okpd2_checked.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
    else:
        st.markdown("""
        <div class="card">
            <div class="card-label">Инструкция</div>
            <div class="card-title">Загрузите рабочий Excel-файл</div>
            <div class="card-sub">
                Обязательна колонка <b>Код ОКПД 2</b>. Все исходные колонки сохраняются.<br>
                В конец добавляются 5 новых колонок:<br>
                • Наименование ОКПД 2 &nbsp;·&nbsp; В Прил. №1 (ДА/НЕТ) &nbsp;·&nbsp; Наим. по Прил. №1<br>
                • В Прил. №2 (ДА/НЕТ) &nbsp;·&nbsp; Наим. по Прил. №2
            </div>
        </div>
        """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
# TAB 2 — ПОИСК ПО КОДУ
# ══════════════════════════════════════════════════════════
with tab2:
    st.markdown("### Поиск по коду ОКПД 2")

    c_in, c_btn = st.columns([5,1])
    with c_in:
        q_code = st.text_input("Код", placeholder="Например: 25.94.11.130 или 13.2",
                               label_visibility="collapsed", key="q_code")
    with c_btn:
        go_code = st.button("Найти", type="primary", key="go_code")

    if go_code and q_code.strip():
        if not bases_ok:
            st.warning("⚠️ Загрузите базы в боковой панели.")
        else:
            raw  = q_code.strip()
            code = clean_code(raw)

            if raw not in st.session_state.hist_code:
                st.session_state.hist_code.insert(0, raw)
                st.session_state.hist_code = st.session_state.hist_code[:10]

            fm, vm = hier_lookup(code, main_d)
            f1, v1 = hier_lookup(code, app1_d)
            f2, v2 = hier_lookup(code, app2_d)

            def badge(found):
                cls = 'badge-yes' if found else 'badge-no'
                txt = '✅ ДА' if found else '❌ НЕТ'
                return f'<span class="card-badge {cls}">{txt}</span>'

            st.markdown(f"""
            <div class="card {'yes' if fm else 'no'}">
                <div class="card-label">📚 Справочник ОКПД 2 {badge(fm)}</div>
                <div class="card-title">{vm if fm else 'Не найдено в справочнике'}</div>
                <div class="card-sub">{code}</div>
            </div>
            <div class="card {'yes' if f1 else 'no'}">
                <div class="card-label">📋 Приложение №1 к ПП 1875 {badge(f1)}</div>
                <div class="card-title">{v1 if f1 else 'Отсутствует в Приложении №1'}</div>
                <div class="card-sub">{code}</div>
            </div>
            <div class="card {'yes' if f2 else 'no'}">
                <div class="card-label">📋 Приложение №2 к ПП 1875 {badge(f2)}</div>
                <div class="card-title">{v2 if f2 else 'Отсутствует в Приложении №2'}</div>
                <div class="card-sub">{code}</div>
            </div>
            """, unsafe_allow_html=True)

    if st.session_state.hist_code:
        st.markdown("---")
        st.markdown("#### 🕘 История")
        for i, h in enumerate(st.session_state.hist_code):
            ca, cb = st.columns([7,1])
            with ca:
                st.markdown(f'<span class="mono">{h}</span>', unsafe_allow_html=True)
            with cb:
                if st.button("🗑️", key=f"dc{i}", type="secondary"):
                    st.session_state.hist_code.pop(i)
                    st.rerun()


# ══════════════════════════════════════════════════════════
# TAB 3 — ПОИСК ПО НАЗВАНИЮ
# ══════════════════════════════════════════════════════════
with tab3:
    st.markdown("### Поиск по наименованию")
    st.caption("Регистр не важен. Поиск подстроки в любом месте наименования.")

    src = st.radio("Искать в:", ["Справочник ОКПД 2", "Приложение №1", "Приложение №2"],
                   horizontal=True, key="name_src")

    c_in2, c_btn2 = st.columns([5,1])
    with c_in2:
        q_name = st.text_input("Наименование или часть слова", placeholder="Например: болт",
                               label_visibility="collapsed", key="q_name")
    with c_btn2:
        go_name = st.button("Найти", type="primary", key="go_name")

    if go_name and q_name.strip():
        if not bases_ok:
            st.warning("⚠️ Загрузите базы в боковой панели.")
        else:
            q  = q_name.strip().lower()
            qs = q_name.strip()

            if qs not in st.session_state.hist_name:
                st.session_state.hist_name.insert(0, qs)
                st.session_state.hist_name = st.session_state.hist_name[:10]

            pool = main_d if src == "Справочник ОКПД 2" else (app1_d if src == "Приложение №1" else app2_d)

            rows = [{'Код ОКПД 2': k, 'Наименование': v}
                    for k,v in pool.items() if q in v.lower()]

            if rows:
                res_df = pd.DataFrame(rows)
                st.success(f"Найдено в «{src}»: **{len(rows)}** совпадений")
                st.dataframe(res_df, use_container_width=True, height=400)

                # Если поиск был по справочнику — автоматически проверяем приложения
                if src == "Справочник ОКПД 2":
                    st.markdown("#### Наличие в Приложениях ПП 1875")
                    check = []
                    for r in rows:
                        c = r['Код ОКПД 2']
                        fa1, _ = hier_lookup(c, app1_d)
                        fa2, _ = hier_lookup(c, app2_d)
                        check.append({
                            'Код ОКПД 2':    c,
                            'Наименование':  pool[c],
                            'В Прил. №1':    '✅ ДА' if fa1 else '❌ НЕТ',
                            'В Прил. №2':    '✅ ДА' if fa2 else '❌ НЕТ',
                        })
                    st.dataframe(pd.DataFrame(check), use_container_width=True)
            else:
                st.error(f"❌ Совпадений не найдено в «{src}».")

    if st.session_state.hist_name:
        st.markdown("---")
        st.markdown("#### 🕘 История")
        for i, h in enumerate(st.session_state.hist_name):
            ca, cb = st.columns([7,1])
            with ca:
                st.markdown(f'<span class="mono">{h}</span>', unsafe_allow_html=True)
            with cb:
                if st.button("🗑️", key=f"dn{i}", type="secondary"):
                    st.session_state.hist_name.pop(i)
                    st.rerun()
