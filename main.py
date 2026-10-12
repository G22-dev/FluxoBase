import streamlit as st
from pathlib import Path
from docxtpl import DocxTemplate
import pandas as pd
import zipfile
import io
import re
import base64
import tempfile
import os
import html

try:
    import pythoncom
    import win32com.client
except ImportError:
    pythoncom = None
    win32com = None

THIS_DIR = Path(__file__).parent
ASSETS_DIR = THIS_DIR / "assets"
LOGO_PATH = ASSETS_DIR / "FluxoBase-Logo.png"

st.set_page_config(
    page_title="FluxoBase",
    page_icon=str(LOGO_PATH) if LOGO_PATH.exists() else "📄",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    [data-testid="stAppViewContainer"] {
        background: linear-gradient(
            145deg, #090d16, #111726, #070a10
        );
        background-attachment: fixed;
    }

    [data-testid="stAppViewContainer"] > .main {
        background: transparent;
    }

    h1, h2, h3, p, span, label {
        color: #f1f5f9;
    }

    [data-testid="stFileUploader"] {
        background-color: #161f30;
        border: 1px dashed #334155;
        border-radius: 8px;
        padding: 10px;
    }

    div[data-testid="stExpander"] {
        background-color: #111827;
        border: 1px solid #334155;
        border-radius: 8px;
    }

    [data-testid="stDownloadButton"] button,
    div.stButton > button {
        width: 100%;
        border-radius: 8px;
        min-height: 42px;
    }

    .preview-container {
        background: #202938;
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 12px;
        margin-top: 12px;
    }

    .preview-description {
        color: #cbd5e1;
        font-size: 14px;
        margin-bottom: 10px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

col1, col2, col3 = st.columns([3, 2, 3])

with col2:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=250)
    else:
        st.warning(
            "Logo não encontrada: assets/FluxoBase-Logo.png"
        )

st.markdown(
    """
    <h1 style="text-align:center; font-weight:bold;">
        Olá, bem-vindo à sua Base de Documentos
    </h1>
    <p style="text-align:center; color:#94a3b8;">
        Gere documentos personalizados diretamente da sua planilha.
    </p>
    """,
    unsafe_allow_html=True,
)

left, center, right = st.columns([1, 2, 1])

with center:
    upload1, upload2 = st.columns(2)

    with upload1:
        st.markdown("**Enviar modelo Word (.docx)**")

        word_document = st.file_uploader(
            "Selecione o modelo Word",
            type=["docx"],
            key="uploader_word",
        )

    with upload2:
        st.markdown("**Enviar planilha Excel (.xlsx)**")

        excel_document = st.file_uploader(
            "Selecione a planilha Excel",
            type=["xlsx"],
            key="uploader_excel",
        )

def safe_filename(value):
    """Remove caracteres inválidos do nome do arquivo."""

    if value is None:
        name = "documento"
    else:
        try:
            name = (
                "documento"
                if pd.isna(value)
                else str(value).strip()
            )
        except (TypeError, ValueError):
            name = str(value).strip()

    name = re.sub(
        r'[<>:"/\\|?*\x00-\x1f]',
        "_",
        name,
    )

    name = name.rstrip(". ")

    return name[:100] or "documento"


def render_document(template_bytes, data):
    """Preenche o modelo Word com os dados de um registro."""

    template = DocxTemplate(io.BytesIO(template_bytes))
    template.render(data)

    output = io.BytesIO()
    template.save(output)

    return output.getvalue()

def convert_docx_to_pdf_with_word(docx_bytes):
    """
    Converte DOCX para PDF pelo Microsoft Word.

    Requer Windows, Microsoft Word e pywin32.
    Retorna os bytes do PDF.
    """

    if os.name != "nt":
        raise RuntimeError(
            "A conversão pelo Microsoft Word exige Windows."
        )

    if pythoncom is None or win32com is None:
        raise RuntimeError(
            "pywin32 não está instalado. Execute: "
            "python -m pip install --upgrade pywin32"
        )

    if not docx_bytes:
        raise ValueError("O documento Word está vazio.")

    word = None
    document = None
    com_initialized = False

    with tempfile.TemporaryDirectory(
        prefix="fluxobase_"
    ) as temp_dir:

        temp_path = Path(temp_dir)
        docx_path = temp_path / "documento.docx"
        pdf_path = temp_path / "documento.pdf"

        docx_path.write_bytes(docx_bytes)

        try:
            pythoncom.CoInitialize()
            com_initialized = True

            word = win32com.client.DispatchEx(
                "Word.Application"
            )

            word.Visible = False
            word.DisplayAlerts = 0

            document = word.Documents.Open(
                FileName=str(docx_path.resolve()),
                ReadOnly=True,
                AddToRecentFiles=False,
                ConfirmConversions=False,
            )

            # Exporta o documento no formato PDF.
            document.ExportAsFixedFormat(
                OutputFileName=str(pdf_path.resolve()),
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0,
            )

            document.Close(SaveChanges=False)
            document = None

            if not pdf_path.is_file():
                raise RuntimeError(
                    "O Word não criou o arquivo PDF."
                )

            pdf_bytes = pdf_path.read_bytes()

            if not pdf_bytes.startswith(b"%PDF-"):
                raise RuntimeError(
                    "O arquivo gerado não é um PDF válido."
                )

            return pdf_bytes

        except Exception as error:
            raise RuntimeError(
                f"{type(error).__name__}: {error}"
            ) from error

        finally:
            if document is not None:
                try:
                    document.Close(SaveChanges=False)
                except Exception:
                    pass

            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass

            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass


def show_pdf_preview(pdf_bytes, filename):

    if not pdf_bytes:
        st.warning("Não há PDF disponível para visualizar.")
        return

    encoded_pdf = base64.b64encode(pdf_bytes).decode("ascii")
    safe_title = html.escape(filename)

    viewer_html = """
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <style>
            * {
                box-sizing: border-box;
            }

            body {
                margin: 0;
                padding: 0;
                background: #111827;
                font-family: Arial, sans-serif;
                color: #f1f5f9;
            }

            .toolbar {
                display: flex;
                align-items: center;
                justify-content: space-between;
                gap: 12px;
                flex-wrap: wrap;
                padding: 12px;
                background: #1e293b;
                border: 1px solid #334155;
                border-radius: 8px 8px 0 0;
            }

            .filename {
                font-size: 13px;
                overflow-wrap: anywhere;
            }

            .open-button {
                display: inline-block;
                background: #2563eb;
                color: white;
                padding: 9px 14px;
                border-radius: 6px;
                text-decoration: none;
                font-size: 13px;
                font-weight: bold;
            }

            .viewer {
                width: 100%;
                height: 780px;
                display: block;
                background: #e5e7eb;
                border: 1px solid #334155;
                border-top: none;
                border-radius: 0 0 8px 8px;
            }

            .message {
                padding: 20px;
                color: #cbd5e1;
                font-size: 14px;
                line-height: 1.6;
            }
        </style>
    </head>

    <body>
        <div class="toolbar">
            <span class="filename">__FILENAME__</span>
            <a id="openPdf" class="open-button"
               href="#" target="_blank" rel="noopener">
                Abrir PDF em nova aba
            </a>
        </div>

        <div id="fallback" class="message">
            Preparando a visualização do documento...
        </div>

        <iframe
            id="pdfFrame"
            class="viewer"
            title="Prévia do documento PDF"
            style="display:none">
        </iframe>

        <script>
        (() => {
            const encoded = "__PDF_BASE64__";
            const fallback = document.getElementById("fallback");
            const frame = document.getElementById("pdfFrame");
            const openLink = document.getElementById("openPdf");

            try {
                const binary = atob(encoded);
                const bytes = new Uint8Array(binary.length);

                for (let i = 0; i < binary.length; i++) {
                    bytes[i] = binary.charCodeAt(i);
                }

                const blob = new Blob(
                    [bytes],
                    { type: "application/pdf" }
                );

                const url = URL.createObjectURL(blob);

                frame.src = url;
                frame.style.display = "block";
                fallback.style.display = "none";
                openLink.href = url;

                // A URL é mantida enquanto o componente estiver aberto.
                // O navegador libera os recursos quando a página é fechada.
            } catch (error) {
                fallback.textContent =
                    "Não foi possível preparar a prévia: " +
                    error.message;
            }
        })();
        </script>
    </body>
    </html>
    """

    viewer_html = (
        viewer_html
        .replace("__PDF_BASE64__", encoded_pdf)
        .replace("__FILENAME__", safe_title)
    )

    st.components.v1.html(
        viewer_html,
        height=850,
        scrolling=True,
    )

def create_zip(df, identifying_column, template_bytes):
    """Gera um DOCX por registro e reúne os arquivos em ZIP."""

    zip_buffer = io.BytesIO()
    errors = []
    generated_count = 0

    with zipfile.ZipFile(
        zip_buffer,
        mode="w",
        compression=zipfile.ZIP_DEFLATED,
    ) as zip_file:

        for index, (_, row) in enumerate(
            df.iterrows(),
            start=1,
        ):
            data = row.to_dict()

            record_name = safe_filename(
                data.get(identifying_column)
            )

            try:
                doc_bytes = render_document(
                    template_bytes,
                    data,
                )

                filename = (
                    f"{index:03d}_{record_name}.docx"
                )

                zip_file.writestr(filename, doc_bytes)
                generated_count += 1

            except Exception as error:
                errors.append(
                    f"Linha {index} ({record_name}): "
                    f"{type(error).__name__}: {error}"
                )

    zip_buffer.seek(0)

    return zip_buffer.getvalue(), errors, generated_count


if word_document is not None and excel_document is not None:

    try:
        template_bytes = word_document.getvalue()

        df = pd.read_excel(
            io.BytesIO(excel_document.getvalue()),
            engine="openpyxl",
        )

        if df.empty:
            st.warning(
                "A planilha não possui registros para processar."
            )
            st.stop()

        if len(df.columns) == 0:
            st.error("A planilha não possui colunas.")
            st.stop()

        # Substitui valores vazios/NaN.
        df = df.astype(object).where(
            pd.notna(df),
            "",
        )

        st.success(
            f"Arquivos carregados! "
            f"{len(df)} registro(s) encontrado(s)."
        )

        with st.expander("Visualizar os dados da planilha"):
            st.dataframe(
                df,
                use_container_width=True,
            )

        identifying_column = df.columns[0]

        st.subheader("Selecionar registro")

        selected_index = st.selectbox(
            "Escolha um registro para visualizar",
            options=list(range(len(df))),
            format_func=lambda i: (
                f"{i + 1} - "
                f"{df.iloc[i][identifying_column]}"
            ),
        )

        line_data = df.iloc[selected_index].to_dict()

        # ----------------------------------
        # GERAÇÃO DA PRÉVIA
        # ----------------------------------

        st.divider()
        st.subheader("Prévia do documento completo")

        st.caption(
            "O documento abaixo corresponde ao registro selecionado, "
            "com os campos do modelo preenchidos."
        )

        preview_bytes = render_document(
            template_bytes,
            line_data,
        )

        pdf_bytes = None
        conversion_error = None

        try:
            with st.spinner("Convertendo documento para PDF..."):
                pdf_bytes = convert_docx_to_pdf_with_word(
                    preview_bytes
                )

        except Exception as error:
            conversion_error = str(error)

        pdf_filename = (
            "Documento_"
            f"{safe_filename(line_data[identifying_column])}"
            ".pdf"
        )

        # ----------------------------------
        # VISUALIZAÇÃO E DOWNLOAD PDF
        # ----------------------------------

        if pdf_bytes:
            st.success("PDF gerado com sucesso.")

            show_pdf_preview(
                pdf_bytes,
                pdf_filename,
            )

            # Um único botão de download da prévia.
            st.download_button(
                label="📑 Baixar prévia em PDF",
                data=pdf_bytes,
                file_name=pdf_filename,
                mime="application/pdf",
                key="download_preview_pdf",
            )

        else:
            st.warning(
                "Não foi possível gerar a prévia em PDF. "
                "Você ainda pode baixar o documento Word."
            )

            if conversion_error:
                with st.expander("Detalhes do erro de conversão"):
                    st.code(
                        conversion_error,
                        language="text",
                    )

        # ----------------------------------
        # DOWNLOAD INDIVIDUAL WORD
        # ----------------------------------

        st.divider()
        st.subheader("Baixar documento individual")

        st.download_button(
            label="📄 Baixar Word deste registro",
            data=preview_bytes,
            file_name=(
                "Documento_"
                f"{safe_filename(line_data[identifying_column])}"
                ".docx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            key="download_individual",
        )

        st.divider()
        st.subheader("Gerar todos os documentos")

        st.write(
            f"Será criado um documento Word para cada um dos "
            f"{len(df)} registros da planilha."
        )

        if st.button(
            "📦 Gerar ZIP com todos os documentos",
            use_container_width=True,
            key="generate_zip",
        ):

            with st.spinner("Gerando os documentos Word..."):
                zip_bytes, errors, generated_count = create_zip(
                    df,
                    identifying_column,
                    template_bytes,
                )

            if errors:
                st.warning(
                    f"{len(errors)} documento(s) "
                    "não puderam ser gerados."
                )

                with st.expander("Ver erros de geração"):
                    for error in errors:
                        st.write(error)

            if generated_count > 0:
                st.success(
                    f"{generated_count} documento(s) "
                    "gerado(s) com sucesso."
                )

                st.download_button(
                    label="⬇️ Baixar ZIP com os documentos Word",
                    data=zip_bytes,
                    file_name="FluxoBase_Documentos.zip",
                    mime="application/zip",
                    key="download_zip",
                )

            else:
                st.error(
                    "Nenhum documento foi gerado. "
                    "Confira o modelo Word e a planilha Excel."
                )

    except Exception as error:
        st.error(
            "Não foi possível processar os arquivos. "
            "Confira o modelo Word e a planilha Excel."
        )

        with st.expander("Detalhes do erro"):
            st.exception(error)

elif word_document is not None or excel_document is not None:

    st.info(
        "Envie o modelo Word e a planilha Excel para continuar."
    )

else:

    st.markdown(
        """
        <div style="
            text-align:center;
            padding:35px;
            background:#111827;
            border:1px solid #334155;
            border-radius:12px;
            margin-top:20px;
        ">
            <h3>Comece por aqui</h3>
            <p>
                Envie seu modelo Word e sua planilha Excel
                para gerar documentos personalizados.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
