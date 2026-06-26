import random
import time
from datetime import timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

try:
    from streamlit_autorefresh import st_autorefresh
    AUTOREFRESH_AVAILABLE = True
except Exception:
    AUTOREFRESH_AVAILABLE = False


# ============================================================
# CONFIGURAZIONE
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

EXCEL_FILENAME = "flashcards_materiali_509_con_esempi_esame.xlsx"
EXCEL_FILE = BASE_DIR / EXCEL_FILENAME

DEFAULT_NUM_QUESTIONS = 60
DEFAULT_DURATION_MINUTES = 30

REQUIRED_COLUMNS = [
    "ID",
    "Argomento",
    "Sottoargomento",
    "Domanda",
    "A",
    "B",
    "C",
    "D",
    "Corretta",
    "Risposta corretta",
    "Spiegazione",
    "Trabocchetto",
]


# ============================================================
# STILE GRAFICO
# ============================================================

def inject_css():
    st.markdown(
        """
        <style>
        html, body, [data-testid="stAppViewContainer"] {
            background:
                radial-gradient(circle at top left, rgba(96, 165, 250, 0.25), transparent 34%),
                radial-gradient(circle at top right, rgba(196, 181, 253, 0.28), transparent 32%),
                linear-gradient(135deg, #f8fafc 0%, #eef2ff 45%, #f8fafc 100%);
            color: #111827;
        }

        [data-testid="stHeader"] {
            background: rgba(255, 255, 255, 0);
        }

        .block-container {
            padding-top: 2.2rem;
            padding-bottom: 4rem;
            max-width: 1180px;
        }

        section[data-testid="stSidebar"] {
            background: rgba(255, 255, 255, 0.60);
            backdrop-filter: blur(26px);
            -webkit-backdrop-filter: blur(26px);
            border-right: 1px solid rgba(255,255,255,0.55);
        }

        h1, h2, h3 {
            letter-spacing: -0.04em;
        }

        .glass-card {
            background: rgba(255, 255, 255, 0.72);
            border: 1px solid rgba(255,255,255,0.65);
            box-shadow: 0 18px 45px rgba(15, 23, 42, 0.10);
            backdrop-filter: blur(26px);
            -webkit-backdrop-filter: blur(26px);
            border-radius: 30px;
            padding: 26px 28px;
            margin: 14px 0 22px 0;
        }

        .glass-card-compact {
            background: rgba(255,255,255,0.62);
            border: 1px solid rgba(255,255,255,0.72);
            box-shadow: 0 8px 24px rgba(15, 23, 42, 0.08);
            backdrop-filter: blur(22px);
            -webkit-backdrop-filter: blur(22px);
            border-radius: 24px;
            padding: 18px 20px;
            margin: 10px 0 16px 0;
        }

        .hero-title {
            font-size: 3rem;
            line-height: 1.02;
            font-weight: 800;
            letter-spacing: -0.06em;
            margin-bottom: 8px;
        }

        .hero-subtitle {
            color: #6b7280;
            font-size: 1.05rem;
            line-height: 1.55;
            max-width: 780px;
        }

        .pill {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 8px 13px;
            border-radius: 999px;
            background: rgba(255,255,255,0.65);
            border: 1px solid rgba(255,255,255,0.75);
            box-shadow: 0 6px 16px rgba(15, 23, 42, 0.06);
            color: #374151;
            font-weight: 600;
            font-size: 0.88rem;
            margin: 4px 6px 4px 0;
        }

        .question-title {
            font-size: 1.28rem;
            font-weight: 750;
            letter-spacing: -0.03em;
            margin-bottom: 6px;
        }

        .question-meta {
            color: #6b7280;
            font-size: 0.92rem;
            margin-bottom: 14px;
        }

        .correct-box {
            background: rgba(22, 163, 74, 0.10);
            border: 1px solid rgba(22, 163, 74, 0.25);
            color: #14532d;
            border-radius: 20px;
            padding: 16px 18px;
            margin-top: 16px;
        }

        .wrong-box {
            background: rgba(220, 38, 38, 0.10);
            border: 1px solid rgba(220, 38, 38, 0.25);
            color: #7f1d1d;
            border-radius: 20px;
            padding: 16px 18px;
            margin-top: 16px;
        }

        .info-box {
            background: rgba(37, 99, 235, 0.09);
            border: 1px solid rgba(37, 99, 235, 0.20);
            color: #1e3a8a;
            border-radius: 20px;
            padding: 16px 18px;
            margin-top: 16px;
        }

        .trap-box {
            background: rgba(245, 158, 11, 0.12);
            border: 1px solid rgba(245, 158, 11, 0.28);
            color: #78350f;
            border-radius: 20px;
            padding: 16px 18px;
            margin-top: 16px;
        }

        .stButton > button {
            border-radius: 999px !important;
            border: 1px solid rgba(255,255,255,0.75) !important;
            background: rgba(255,255,255,0.78) !important;
            box-shadow: 0 8px 22px rgba(15, 23, 42, 0.08) !important;
            color: #111827 !important;
            font-weight: 650 !important;
            padding: 0.70rem 1.15rem !important;
        }

        .stButton > button[kind="primary"] {
            background: linear-gradient(135deg, #2563eb, #60a5fa) !important;
            color: white !important;
            border: 1px solid rgba(255,255,255,0.45) !important;
        }

        label[data-baseweb="radio"] {
            background: rgba(255,255,255,0.64);
            border: 1px solid rgba(255,255,255,0.72);
            border-radius: 18px;
            padding: 12px 14px;
            margin-bottom: 8px;
            box-shadow: 0 6px 16px rgba(15, 23, 42, 0.045);
        }

        label[data-baseweb="radio"]:hover {
            border-color: rgba(37, 99, 235, 0.35);
            background: rgba(255,255,255,0.84);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# CARICAMENTO DATI
# ============================================================

@st.cache_data
def load_flashcards_from_file(file_source) -> pd.DataFrame:
    df = pd.read_excel(file_source)

    missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Mancano queste colonne nel file Excel: {missing_cols}")

    df = df.dropna(subset=["Domanda", "A", "B", "C", "D", "Corretta"]).copy()
    df["Corretta"] = df["Corretta"].astype(str).str.strip().str.upper()

    for col in REQUIRED_COLUMNS:
        df[col] = df[col].fillna("").astype(str)

    df = df[df["Corretta"].isin(["A", "B", "C", "D"])].copy()

    return df.reset_index(drop=True)


def get_dataframe():
    """
    Prima prova a caricare l'Excel dalla stessa cartella dello script.
    Se non lo trova, mostra un uploader nella web app.
    """
    if EXCEL_FILE.exists():
        return load_flashcards_from_file(EXCEL_FILE)

    st.error("Non trovo il file Excel nella cartella dello script.")
    st.write("Percorso cercato:")
    st.code(str(EXCEL_FILE))

    st.info(
        "Soluzione veloce: carica qui sotto il file Excel, oppure mettilo nella stessa cartella dello script."
    )

    uploaded_file = st.file_uploader(
        "Carica il file Excel delle flashcard",
        type=["xlsx"],
    )

    if uploaded_file is None:
        st.stop()

    return load_flashcards_from_file(uploaded_file)


# ============================================================
# FUNZIONI DOMANDE
# ============================================================

def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def format_time(seconds: float) -> str:
    seconds = max(0, int(seconds))
    return str(timedelta(seconds=seconds))


def make_question_from_row(row: pd.Series) -> dict:
    correct_letter = clean_text(row["Corretta"]).upper()

    options = [
        {"original_letter": "A", "text": clean_text(row["A"])},
        {"original_letter": "B", "text": clean_text(row["B"])},
        {"original_letter": "C", "text": clean_text(row["C"])},
        {"original_letter": "D", "text": clean_text(row["D"])},
    ]

    random.shuffle(options)

    correct_index = 0
    for i, option in enumerate(options):
        if option["original_letter"] == correct_letter:
            correct_index = i
            break

    return {
        "id": clean_text(row["ID"]),
        "argomento": clean_text(row["Argomento"]),
        "sottoargomento": clean_text(row["Sottoargomento"]),
        "domanda": clean_text(row["Domanda"]),
        "options": options,
        "correct_index": correct_index,
        "spiegazione": clean_text(row["Spiegazione"]),
        "trabocchetto": clean_text(row["Trabocchetto"]),
    }


def create_exam(df: pd.DataFrame, num_questions: int) -> list:
    num_questions = min(num_questions, len(df))
    sampled = df.sample(n=num_questions).reset_index(drop=True)
    return [make_question_from_row(row) for _, row in sampled.iterrows()]


def create_training_question(df: pd.DataFrame, argomento: str, sottoargomento: str) -> dict:
    filtered = df.copy()

    if argomento != "Tutti gli argomenti":
        filtered = filtered[filtered["Argomento"] == argomento]

    if sottoargomento != "Tutti i sottoargomenti":
        filtered = filtered[filtered["Sottoargomento"] == sottoargomento]

    if filtered.empty:
        filtered = df.copy()

    row = filtered.sample(n=1).iloc[0]
    return make_question_from_row(row)


def calculate_exam_results(questions: list, answers: dict) -> tuple[int, list]:
    score = 0
    results = []

    for i, question in enumerate(questions):
        selected_index = answers.get(i)
        correct_index = question["correct_index"]

        is_correct = selected_index == correct_index

        if is_correct:
            score += 1

        selected_text = (
            question["options"][selected_index]["text"]
            if selected_index is not None
            else "Non risposta"
        )

        correct_text = question["options"][correct_index]["text"]

        results.append(
            {
                "N": i + 1,
                "ID": question["id"],
                "Argomento": question["argomento"],
                "Sottoargomento": question["sottoargomento"],
                "Domanda": question["domanda"],
                "Tua risposta": selected_text,
                "Risposta corretta": correct_text,
                "Esito": "Corretta" if is_correct else "Sbagliata",
                "Spiegazione": question["spiegazione"],
                "Trabocchetto": question["trabocchetto"],
            }
        )

    return score, results


# ============================================================
# RESET
# ============================================================

def reset_exam_session():
    keys = [
        "exam_started",
        "exam_submitted",
        "exam_questions",
        "exam_answers",
        "exam_start_time",
        "exam_duration_seconds",
    ]

    for key in keys:
        if key in st.session_state:
            del st.session_state[key]

    for key in list(st.session_state.keys()):
        if key.startswith("exam_q_"):
            del st.session_state[key]


def reset_training_session():
    keys = [
        "training_question",
        "training_answered",
        "training_selected_index",
        "training_correct_count",
        "training_wrong_count",
        "training_total_count",
    ]

    for key in keys:
        if key in st.session_state:
            del st.session_state[key]

    for key in list(st.session_state.keys()):
        if key.startswith("training_radio"):
            del st.session_state[key]


# ============================================================
# COMPONENTI GRAFICI
# ============================================================

def glass_card_start():
    st.markdown('<div class="glass-card">', unsafe_allow_html=True)


def glass_card_end():
    st.markdown("</div>", unsafe_allow_html=True)


def render_question_card(question: dict, number=None):
    title = f"Domanda {number}" if number is not None else "Domanda"

    st.markdown(
        f"""
        <div class="question-title">{title}</div>
        <div class="question-meta">
            {question["argomento"]} · {question["sottoargomento"]}
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.write(question["domanda"])


def render_feedback(question: dict, selected_index):
    correct_index = question["correct_index"]
    correct_text = question["options"][correct_index]["text"]

    if selected_index is None:
        st.markdown(
            """
            <div class="wrong-box">
            <b>Risposta non data.</b>
            </div>
            """,
            unsafe_allow_html=True,
        )
    elif selected_index == correct_index:
        st.markdown(
            """
            <div class="correct-box">
            <b>Risposta corretta.</b>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        selected_text = question["options"][selected_index]["text"]
        st.markdown(
            f"""
            <div class="wrong-box">
            <b>Risposta sbagliata.</b><br>
            La tua risposta: {selected_text}
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown(
        f"""
        <div class="info-box">
        <b>Risposta corretta:</b><br>
        {correct_text}
        </div>
        """,
        unsafe_allow_html=True,
    )

    if question["spiegazione"]:
        st.markdown(
            f"""
            <div class="info-box">
            <b>Perché:</b><br>
            {question["spiegazione"]}
            </div>
            """,
            unsafe_allow_html=True,
        )

    if question["trabocchetto"]:
        st.markdown(
            f"""
            <div class="trap-box">
            <b>Trabocchetto:</b><br>
            {question["trabocchetto"]}
            </div>
            """,
            unsafe_allow_html=True,
        )


# ============================================================
# AVVIO APP
# ============================================================

st.set_page_config(
    page_title="Materiali Exam Lab",
    page_icon="🧱",
    layout="wide",
)

inject_css()

df = get_dataframe()


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.markdown("## 🧱 Materiali")
    st.caption("Exam Lab")

    st.markdown("---")

    page = st.radio(
        "Modalità",
        [
            "🏠 Dashboard",
            "📝 Simulazione esame",
            "🎯 Allenamento",
            "📊 Database",
        ],
        label_visibility="collapsed",
    )

    st.markdown("---")

    st.markdown(
        f"""
        <div class="glass-card-compact">
            <b>{len(df)}</b><br>
            domande disponibili
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# DASHBOARD
# ============================================================

if page == "🏠 Dashboard":
    st.markdown(
        """
        <div class="glass-card">
            <div class="hero-title">Materiali Exam Lab</div>
            <div class="hero-subtitle">
                Allenamento e simulazione d’esame per Materiali da costruzione,
                innovazione e sostenibilità.
            </div>
            <div style="margin-top: 18px;">
                <span class="pill">📝 Esame 60 domande</span>
                <span class="pill">⏱️ Timer 30 minuti</span>
                <span class="pill">🎯 Allenamento per categoria</span>
                <span class="pill">🔀 Risposte rimescolate</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("Domande totali", len(df))

    with col2:
        st.metric("Argomenti", df["Argomento"].nunique())

    with col3:
        st.metric("Sottoargomenti", df["Sottoargomento"].nunique())

    st.subheader("Distribuzione per argomento")

    topic_counts = df["Argomento"].value_counts().reset_index()
    topic_counts.columns = ["Argomento", "Numero domande"]

    st.dataframe(topic_counts, use_container_width=True, hide_index=True)


# ============================================================
# SIMULAZIONE ESAME
# ============================================================

elif page == "📝 Simulazione esame":
    st.markdown(
        """
        <div class="glass-card">
            <div class="hero-title">Simulazione esame</div>
            <div class="hero-subtitle">
                60 domande casuali, 4 risposte rimescolate e timer da 30 minuti.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if "exam_started" not in st.session_state:
        st.session_state.exam_started = False

    if "exam_submitted" not in st.session_state:
        st.session_state.exam_submitted = False

    if "exam_answers" not in st.session_state:
        st.session_state.exam_answers = {}

    if not st.session_state.exam_started:
        col1, col2, col3 = st.columns(3)

        with col1:
            num_questions = st.number_input(
                "Numero domande",
                min_value=1,
                max_value=len(df),
                value=min(DEFAULT_NUM_QUESTIONS, len(df)),
                step=1,
            )

        with col2:
            duration_minutes = st.number_input(
                "Durata in minuti",
                min_value=1,
                max_value=180,
                value=DEFAULT_DURATION_MINUTES,
                step=1,
            )

        with col3:
            st.write("")
            st.write("")
            if st.button("▶️ Avvia esame", type="primary", use_container_width=True):
                st.session_state.exam_started = True
                st.session_state.exam_submitted = False
                st.session_state.exam_questions = create_exam(df, int(num_questions))
                st.session_state.exam_answers = {}
                st.session_state.exam_start_time = time.time()
                st.session_state.exam_duration_seconds = int(duration_minutes * 60)

                for key in list(st.session_state.keys()):
                    if key.startswith("exam_q_"):
                        del st.session_state[key]

                st.rerun()

    else:
        questions = st.session_state.exam_questions

        elapsed = time.time() - st.session_state.exam_start_time
        remaining = st.session_state.exam_duration_seconds - elapsed

        if remaining <= 0 and not st.session_state.exam_submitted:
            st.session_state.exam_submitted = True
            remaining = 0

        if AUTOREFRESH_AVAILABLE and not st.session_state.exam_submitted:
            st_autorefresh(interval=5000, key="exam_timer_refresh")

        answered_count = sum(
            1 for value in st.session_state.exam_answers.values()
            if value is not None
        )

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric("Tempo rimasto", format_time(remaining))

        with col2:
            st.metric("Risposte date", f"{answered_count}/{len(questions)}")

        with col3:
            st.metric("Stato", "Terminato" if st.session_state.exam_submitted else "In corso")

        with col4:
            if st.button("🔄 Nuova prova", use_container_width=True):
                reset_exam_session()
                st.rerun()

        if not st.session_state.exam_submitted:
            for i, question in enumerate(questions):
                glass_card_start()

                render_question_card(question, i + 1)

                displayed_letters = ["A", "B", "C", "D"]
                option_labels = [
                    f"{displayed_letters[j]}) {option['text']}"
                    for j, option in enumerate(question["options"])
                ]

                selected = st.radio(
                    "Scegli una risposta:",
                    options=option_labels,
                    index=None,
                    key=f"exam_q_{i}",
                )

                if selected is not None:
                    st.session_state.exam_answers[i] = option_labels.index(selected)
                else:
                    st.session_state.exam_answers[i] = None

                glass_card_end()

            if st.button("✅ Consegna esame", type="primary", use_container_width=True):
                st.session_state.exam_submitted = True
                st.rerun()

        else:
            score, results = calculate_exam_results(
                st.session_state.exam_questions,
                st.session_state.exam_answers,
            )

            total = len(st.session_state.exam_questions)
            percentage = score / total * 100
            wrong = total - score

            col1, col2, col3 = st.columns(3)

            with col1:
                st.metric("Punteggio", f"{score}/{total}")

            with col2:
                st.metric("Percentuale", f"{percentage:.1f}%")

            with col3:
                st.metric("Errori", wrong)

            results_df = pd.DataFrame(results)

            st.subheader("Riepilogo per argomento")

            topic_summary = (
                results_df
                .assign(Corretta_num=lambda x: x["Esito"].eq("Corretta").astype(int))
                .groupby("Argomento")
                .agg(
                    Domande=("Esito", "count"),
                    Corrette=("Corretta_num", "sum"),
                )
                .reset_index()
            )

            topic_summary["Errori"] = topic_summary["Domande"] - topic_summary["Corrette"]
            topic_summary["Percentuale"] = (
                topic_summary["Corrette"] / topic_summary["Domande"] * 100
            ).round(1)

            st.dataframe(topic_summary, use_container_width=True, hide_index=True)

            st.subheader("Domande sbagliate")

            wrong_results = results_df[results_df["Esito"] == "Sbagliata"]

            if wrong_results.empty:
                st.success("Hai risposto correttamente a tutte le domande.")
            else:
                for _, row in wrong_results.iterrows():
                    with st.expander(
                        f"Domanda {row['N']} — {row['Argomento']} / {row['Sottoargomento']}"
                    ):
                        st.write("**Domanda:**")
                        st.write(row["Domanda"])

                        st.write("**La tua risposta:**")
                        st.error(row["Tua risposta"])

                        st.write("**Risposta corretta:**")
                        st.success(row["Risposta corretta"])

                        st.write("**Spiegazione:**")
                        st.write(row["Spiegazione"])

                        st.write("**Trabocchetto:**")
                        st.warning(row["Trabocchetto"])

            st.subheader("Tabella completa risultati")

            st.dataframe(results_df, use_container_width=True, hide_index=True)

            csv = results_df.to_csv(index=False).encode("utf-8-sig")

            st.download_button(
                label="⬇️ Scarica risultati CSV",
                data=csv,
                file_name="risultati_simulazione_materiali.csv",
                mime="text/csv",
                use_container_width=True,
            )


# ============================================================
# ALLENAMENTO
# ============================================================

elif page == "🎯 Allenamento":
    st.markdown(
        """
        <div class="glass-card">
            <div class="hero-title">Allenamento</div>
            <div class="hero-subtitle">
                Scegli una categoria, rispondi e ricevi subito il feedback.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if "training_correct_count" not in st.session_state:
        st.session_state.training_correct_count = 0

    if "training_wrong_count" not in st.session_state:
        st.session_state.training_wrong_count = 0

    if "training_total_count" not in st.session_state:
        st.session_state.training_total_count = 0

    if "training_answered" not in st.session_state:
        st.session_state.training_answered = False

    if "training_selected_index" not in st.session_state:
        st.session_state.training_selected_index = None

    argomenti = ["Tutti gli argomenti"] + sorted(df["Argomento"].dropna().unique().tolist())

    col1, col2, col3 = st.columns([1.2, 1.2, 0.8])

    with col1:
        selected_argomento = st.selectbox("Argomento", argomenti)

    filtered_for_sub = df.copy()

    if selected_argomento != "Tutti gli argomenti":
        filtered_for_sub = filtered_for_sub[filtered_for_sub["Argomento"] == selected_argomento]

    sottoargomenti = ["Tutti i sottoargomenti"] + sorted(
        filtered_for_sub["Sottoargomento"].dropna().unique().tolist()
    )

    with col2:
        selected_sottoargomento = st.selectbox("Sottoargomento", sottoargomenti)

    with col3:
        st.write("")
        st.write("")
        if st.button("Reset score", use_container_width=True):
            reset_training_session()
            st.rerun()

    col_a, col_b, col_c = st.columns(3)

    with col_a:
        st.metric("Corrette", st.session_state.training_correct_count)

    with col_b:
        st.metric("Sbagliate", st.session_state.training_wrong_count)

    with col_c:
        total_training = st.session_state.training_total_count
        accuracy = (
            st.session_state.training_correct_count / total_training * 100
            if total_training > 0
            else 0
        )
        st.metric("Accuracy", f"{accuracy:.1f}%")

    if "training_question" not in st.session_state:
        st.session_state.training_question = create_training_question(
            df,
            selected_argomento,
            selected_sottoargomento,
        )
        st.session_state.training_answered = False
        st.session_state.training_selected_index = None

    question = st.session_state.training_question

    glass_card_start()

    render_question_card(question)

    displayed_letters = ["A", "B", "C", "D"]
    option_labels = [
        f"{displayed_letters[j]}) {option['text']}"
        for j, option in enumerate(question["options"])
    ]

    selected = st.radio(
        "Scegli una risposta:",
        options=option_labels,
        index=None,
        key="training_radio_current",
        disabled=st.session_state.training_answered,
    )

    if selected is not None and not st.session_state.training_answered:
        st.session_state.training_selected_index = option_labels.index(selected)

    col_verify, col_next = st.columns(2)

    with col_verify:
        if st.button(
            "Verifica risposta",
            type="primary",
            use_container_width=True,
            disabled=st.session_state.training_answered,
        ):
            if st.session_state.training_selected_index is None:
                st.warning("Seleziona una risposta prima di verificare.")
            else:
                st.session_state.training_answered = True
                st.session_state.training_total_count += 1

                if st.session_state.training_selected_index == question["correct_index"]:
                    st.session_state.training_correct_count += 1
                else:
                    st.session_state.training_wrong_count += 1

                st.rerun()

    with col_next:
        if st.button("Prossima domanda", use_container_width=True):
            for key in list(st.session_state.keys()):
                if key.startswith("training_radio"):
                    del st.session_state[key]

            st.session_state.training_question = create_training_question(
                df,
                selected_argomento,
                selected_sottoargomento,
            )
            st.session_state.training_answered = False
            st.session_state.training_selected_index = None
            st.rerun()

    if st.session_state.training_answered:
        render_feedback(question, st.session_state.training_selected_index)

    glass_card_end()


# ============================================================
# DATABASE
# ============================================================

elif page == "📊 Database":
    st.markdown(
        """
        <div class="glass-card">
            <div class="hero-title">Database domande</div>
            <div class="hero-subtitle">
                Controlla e cerca tutte le domande presenti nel file Excel.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("Domande totali", len(df))

    with col2:
        st.metric("Argomenti", df["Argomento"].nunique())

    with col3:
        st.metric("Sottoargomenti", df["Sottoargomento"].nunique())

    st.subheader("Domande per argomento")

    topic_counts = df["Argomento"].value_counts().reset_index()
    topic_counts.columns = ["Argomento", "Numero domande"]

    st.dataframe(topic_counts, use_container_width=True, hide_index=True)

    st.subheader("Database completo")

    search = st.text_input("Cerca nel testo delle domande")

    shown_df = df.copy()

    if search.strip():
        shown_df = shown_df[
            shown_df["Domanda"].str.contains(search.strip(), case=False, na=False)
        ]

    st.dataframe(
        shown_df[
            [
                "ID",
                "Argomento",
                "Sottoargomento",
                "Domanda",
                "A",
                "B",
                "C",
                "D",
                "Corretta",
                "Spiegazione",
                "Trabocchetto",
            ]
        ],
        use_container_width=True,
        hide_index=True,
    )