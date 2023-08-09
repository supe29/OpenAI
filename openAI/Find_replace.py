from pathlib import Path
import pythoncom
import win32com.client
import openai
import streamlit as st

# Set API key
openai.api_key = st.secrets['pass']

st.header("Arjun application")
st.subheader("Upload Microsoft Word Document")
doc_files = st.file_uploader("Upload .docx file", type=[
                             "docx"], accept_multiple_files=True)

article_text = st.text_area("Please enter your code")

temp = st.slider("Temperature", 0.0, 1.0, 0.5)

if st.button("Generate"):
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt="Plz write summary in short:" + article_text,
        max_tokens=3000,
        temperature=temp
    )
    res = response["choices"][0]["text"]
    st.info(res)

    # Word document replacement
    current_dir = Path(__file__).resolve().parent
    output_dir = current_dir / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    # Find and replace
    find_str = "2022"
    replace_with = res
    wd_replace = 2
    wd_find_wrap = 1

    # Open Word
    pythoncom.CoInitialize()
    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = False
    word_app.DisplayAlerts = False

    for doc_file in doc_files:
        doc_path = current_dir / doc_file.name  # Get the path of the uploaded file
        with open(doc_path, "wb") as f:
            f.write(doc_file.getbuffer() if hasattr(
                doc_file, "getbuffer") else doc_file.read())

        word_app.Documents.Open(str(doc_path))

        word_app.Selection.Find.Execute(
            FindText=find_str,
            ReplaceWith=replace_with,
            Replace=wd_replace,
            Forward=True,
            MatchCase=True,
            MatchWholeWord=False,
            MatchWildcards=True,
            MatchSoundsLike=False,
            MatchAllWordForms=False,
            Wrap=wd_find_wrap,
            Format=True
        )

        for i in range(word_app.ActiveDocument.Shapes.Count):
            if word_app.ActiveDocument.Shapes(i + 1).TextFrame.HasText:
                words = word_app.ActiveDocument.Shapes(
                    i + 1).TextFrame.TextRange.Words
                for j in range(words.Count):
                    if word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text == find_str:
                        word_app.ActiveDocument.Shapes(
                            i + 1).TextFrame.TextRange.Words.Item(j + 1).Text = replace_with

        # Saving new file
        output_path = output_dir / \
            f"{doc_file.name.stem}_replaced{doc_file.name.suffix}"
        word_app.ActiveDocument.SaveAs(str(output_path))
        word_app.ActiveDocument.Close(SaveChanges=False)

    word_app.Application.Quit()
    pythoncom.CoUninitialize()
