import os
import pypandoc
from pypandoc.pandoc_download import download_pandoc

class WordMark:
    def __init__(self):
        download_pandoc()
        output_folder = "../data/output"
        self.output_folder_word = os.path.join(output_folder, "word")
        self.output_folder_markdown = os.path.join(output_folder, "markdown")
        self.input_folder = "../data/input"

        # Erstelle die Ordner, falls sie nicht existieren
        os.makedirs(self.input_folder, exist_ok=True)
        os.makedirs(self.output_folder_word, exist_ok=True)
        os.makedirs(self.output_folder_markdown, exist_ok=True)

        print(f"Output folder for Word files: {self.output_folder_word}")
        print(f"Output folder for Markdown files: {self.output_folder_markdown}")
        print(f"Input folder: {self.input_folder}")

    def md_to_docx_pandoc(self, md_file=None):
        if md_file:
            files = [md_file]
        else:
            files = [f for f in os.listdir(self.input_folder) if f.lower().endswith('.md')]
        for file in files:
            file_path = os.path.join(self.input_folder, file)
            file_name = os.path.splitext(file)[0]
            output_file = os.path.join(self.output_folder_word, file_name + '.docx')
            print(f"Converting {file_path} to Word file {output_file}")
            pypandoc.convert_file(source_file=file_path, to='docx', outputfile=output_file)
            print(f"Converted {file_path} → {output_file}")

    def word_to_md_pandoc(self, word_file=None):
        if word_file:
            files = [word_file]
        else:
            files = [f for f in os.listdir(self.input_folder) if f.lower().endswith('.docx')]
        for file in files:
            file_path = os.path.join(self.input_folder, file)
            file_name = os.path.splitext(file)[0]
            output_file = os.path.join(self.output_folder_markdown, file_name + '.md')
            print(f"Converting {file_path} to Markdown file {output_file}")
            pypandoc.convert_file(source_file=file_path, to='md', outputfile=output_file)
            print(f"Converted {file_path} → {output_file}")

if __name__ == "__main__":
    wordmark = WordMark()
    wordmark.md_to_docx_pandoc()   # Batch-Verarbeitung aller .md-Dateien
    wordmark.word_to_md_pandoc()   # Batch-Verarbeitung aller .docx-Dateien
