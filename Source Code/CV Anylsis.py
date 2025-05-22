import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from fpdf import FPDF
import os


def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            text += page.extract_text() or ""
    except Exception as e:
        messagebox.showerror("Error", f"Could not read file {pdf_path}\n{str(e)}")
    return text


def calculate_score(text):
    experience_score = 5 if "years" in text else 0
    skills = ["Python", "Machine Learning", "Data Analysis", "Communication"]
    skill_score = sum(1 for skill in skills if skill in text) * 3
    certification_score = text.lower().count("certification") * 2
    total_score = experience_score + skill_score + certification_score
    return experience_score, skill_score, certification_score, total_score


class CVAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CV Analyzer")
        self.root.attributes('-fullscreen', True)
        self.root.config(bg="#2b2b2b")
        self.root.bind("<Escape>", lambda e: root.quit())

        title_frame = tk.Frame(root, bg="#34495e", pady=10)
        title_frame.pack(fill="x", padx=5, pady=5)
        title_label = tk.Label(title_frame, text="CV Analyzer System",
                               font=("Arial", 16, "bold"), fg="#ecf0f1", bg="#34495e")
        title_label.pack(pady=10)

        file_frame = tk.Frame(root, bg="#2b2b2b")
        file_frame.pack(pady=10)
        select_button = tk.Button(file_frame, text="Select CV Files", command=self.select_files,
                                  font=("Arial", 12), bg="#3498db", fg="white", padx=20, pady=5)
        select_button.grid(row=0, column=0, padx=10)

        self.analyze_button = tk.Button(file_frame, text="Analyze CVs", command=self.analyze_cvs,
                                        font=("Arial", 12), bg="#1abc9c", fg="white", padx=20, pady=5)
        self.analyze_button.config(state=tk.DISABLED)
        self.analyze_button.grid(row=0, column=1, padx=10)

        self.export_pdf_button = tk.Button(file_frame, text="Export to PDF", command=self.export_to_pdf,
                                            font=("Arial", 12), bg="#e67e22", fg="white", padx=20, pady=5)
        self.export_pdf_button.config(state=tk.DISABLED)
        self.export_pdf_button.grid(row=0, column=2, padx=10)

        self.export_excel_button = tk.Button(file_frame, text="Export to Excel", command=self.export_to_excel,
                                              font=("Arial", 12), bg="#e67e22", fg="white", padx=20, pady=5)
        self.export_excel_button.config(state=tk.DISABLED)
        self.export_excel_button.grid(row=0, column=3, padx=10)

        self.result_frame = tk.Frame(root, bg="#2b2b2b", pady=20)
        self.result_frame.pack(fill="both", expand=True, padx=5, pady=5)

        self.cv_files = []
        self.score_threshold = 10  

    def select_files(self):
        self.cv_files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if len(self.cv_files) != 2:
            messagebox.showerror("Error", "Please select exactly two PDF files.")
            self.cv_files = []
        else:
            self.analyze_button.config(state=tk.NORMAL)

    def analyze_cvs(self):
        cv_data = []
        for file_path in self.cv_files:
            text = extract_text_from_pdf(file_path)
            exp_score, skill_score, cert_score, total_score = calculate_score(text)
            status = "Passed" if total_score >= self.score_threshold else "Rejected"
            cv_data.append({
                "File": file_path.split("/")[-1],
                "Experience Score": exp_score,
                "Skill Score": skill_score,
                "Certification Score": cert_score,
                "Total Score": total_score,
                "Status": status
            })

        self.df = pd.DataFrame(cv_data)
        self.display_results(self.df)
        self.export_pdf_button.config(state=tk.NORMAL)
        self.export_excel_button.config(state=tk.NORMAL)

    def display_results(self, df):
        for widget in self.result_frame.winfo_children():
            widget.destroy()

        header = ["File", "Experience Score", "Skill Score", "Certification Score", "Total Score", "Status"]
        for col, text in enumerate(header):
            label = tk.Label(self.result_frame, text=text, font=("Arial", 12, "bold"), fg="#ecf0f1", bg="#2b2b2b")
            label.grid(row=0, column=col, padx=10, pady=5)

        for i, row in df.iterrows():
            for j, value in enumerate(row):
                label = tk.Label(self.result_frame, text=value, font=("Arial", 10), fg="#ecf0f1", bg="#2b2b2b")
                label.grid(row=i + 1, column=j, padx=10, pady=5)

        if df["Total Score"].min() == 0:
            best_cv_label = tk.Label(self.result_frame, text="Both candidates are below the required threshold.",
                                     font=("Arial", 14, "bold"), fg="#e74c3c", bg="#2b2b2b", pady=20)
        elif df["Total Score"].nunique() == 1:
            best_cv_label = tk.Label(self.result_frame, text="Both candidates have the same score.",
                                     font=("Arial", 14, "bold"), fg="#f1c40f", bg="#2b2b2b", pady=20)
        else:
            best_cv = df.loc[df["Total Score"].idxmax()]["File"]
            best_cv_label = tk.Label(self.result_frame, text=f"Best Candidate: {best_cv}", font=("Arial", 14, "bold"),
                                     fg="#f1c40f", bg="#2b2b2b", pady=20)
        best_cv_label.grid(row=len(df) + 2, column=0, columnspan=len(header), pady=10)

     
        fig, axs = plt.subplots(1, 3, figsize=(15, 4), facecolor='#2b2b2b')

        for i, row in df.iterrows():
            ax = axs[i]
            ax.bar(["Experience", "Skills", "Certification"],
                   [row["Experience Score"], row["Skill Score"], row["Certification Score"]],
                   color=['#3498db', '#1abc9c', '#e74c3c'])
            ax.set_title(f"{row['File']} - Detail Score", color="white")
            ax.set_ylim(0, max(df["Total Score"]) + 5)
            ax.tick_params(colors='white')
            ax.set_facecolor("#2b2b2b")
            for spine in ax.spines.values():
                spine.set_edgecolor("white")

        ax = axs[2]
        ax.bar(df["File"], df["Total Score"], color=['#3498db', '#1abc9c'])
        ax.set_title("Total Score Comparison", color="white")
        ax.set_ylim(0, max(df["Total Score"]) + 5)
        ax.tick_params(colors='white')
        ax.set_facecolor("#2b2b2b")
        for spine in ax.spines.values():
            spine.set_edgecolor("white")

        plt.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=self.result_frame)
        canvas.draw()
        canvas.get_tk_widget().grid(row=len(df) + 4, column=0, columnspan=len(header), pady=20)

        
        self.graph_file = "cv_analysis_graphs.png"
        fig.savefig(self.graph_file)

    def export_to_pdf(self):
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, 'CV Analysis Report', ln=True, align='C')
        pdf.set_font("Arial", size=12)

       
        summary = "Summary of Insights:\n"
        summary += "Most candidates excelled in technical skills but lacked certifications.\n"
        pdf.multi_cell(0, 10, summary)

       
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 10, 'Detailed Results:', ln=True)
        pdf.set_font("Arial", size=12)
        for i, row in self.df.iterrows():
            pdf.cell(0, 10, f"File: {row['File']}, Total Score: {row['Total Score']}, Status: {row['Status']}", ln=True)
        pdf.cell(0, 10, '', ln=True)

        pdf.image(self.graph_file, x=10, y=None, w=180)
        pdf.output("cv_analysis_report.pdf")
        messagebox.showinfo("Success", "Report has been exported to PDF successfully.")

    def export_to_excel(self):
        self.df.to_excel("cv_analysis_results.xlsx", index=False)
        messagebox.showinfo("Success", "Results have been exported to Excel successfully.")


if __name__ == "__main__":
    root = tk.Tk()
    app = CVAnalyzerApp(root)
    root.mainloop()
