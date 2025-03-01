import pypandoc
from pypandoc.pandoc_download import download_pandoc

class WordMark:
    def __init__(self):
        # Installiere pandoc, falls nicht vorhanden
        download_pandoc()
        
        output_folder = "../data/output"
        self.output_folder_word = f"{output_folder}/word"
        self.output_folder_markdown = f"{output_folder}/markdown"
        self.input_folder = "../data/input"
        print(f"Output folder for Word files: {self.output_folder_word}")
        print(f"Output folder for Markdown files: {self.output_folder_markdown}")
        print(f"Input folder: {self.input_folder}")

    def md_to_docx_pandoc(self, md_file):
        file_path_md_file = f"{self.input_folder}/{md_file}"
        print(f"Converting {file_path_md_file} to Word file")
        file_name = md_file.split(".")[0]
        print(f"File name: {file_name}")
        file_path_docx_file = f"{self.output_folder_word}/{file_name}.docx"
        print(f"Output file path: {file_path_docx_file}")

        output = pypandoc.convert_file(source_file=file_path_md_file, to='docx', outputfile=file_path_docx_file)
        print(f"Converted {file_path_md_file} → {file_path_docx_file}")

    def word_to_md_pandoc(self, word_file):
        file_path_word_file = f"{self.input_folder}/{word_file}"
        print(f"Converting {file_path_word_file} to Markdown file")
        file_name = word_file.split(".")[0]
        print(f"File name: {file_name}")
        file_path_md_file = f"{self.output_folder_markdown}/{file_name}.md"
        print(f"Output file path: {file_path_md_file}")

        output = pypandoc.convert_file(source_file=file_path_word_file, to='md', outputfile=file_path_md_file)
        print(f"Converted {file_path_word_file} → {file_path_md_file}")

if __name__ == "__main__":
    wordmark = WordMark()
    
    wordmark.word_to_md_pandoc("Word-Test.docx")
    wordmark.md_to_docx_pandoc("Markdown-Test.md")
