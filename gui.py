import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


def select_source_dir():
    dir_path = filedialog.askdirectory(title="Select Source Directory")
    if dir_path:
        source_entry.delete(0, tk.END)
        source_entry.insert(0, dir_path)


def select_target_dir():
    dir_path = filedialog.askdirectory(title="Select Target Directory")
    if dir_path:
        target_entry.delete(0, tk.END)
        target_entry.insert(0, dir_path)


def process_files():
    from docx_converter import docx_to_markdown
    from pptx_converter import pptx_to_markdown

    source_dir = source_entry.get().strip()
    target_dir = target_entry.get().strip()

    if not source_dir:
        messagebox.showwarning("Warning", "Please select source directory")
        return

    if not target_dir:
        messagebox.showwarning("Warning", "Please select target directory")
        return

    if not os.path.exists(source_dir):
        messagebox.showerror("Error", f"Source directory does not exist: {source_dir}")
        return

    btn_process.config(state="disabled")
    status_label.config(text="Scanning and processing files...")
    result_text.delete(1.0, tk.END)

    total_files = 0
    success_count = 0
    fail_count = 0

    try:
        for root_dir, dirs, files in os.walk(source_dir):
            rel_path = os.path.relpath(root_dir, source_dir)

            if rel_path == '.':
                target_subdir = target_dir
            else:
                target_subdir = os.path.join(target_dir, rel_path)

            if not os.path.exists(target_subdir):
                os.makedirs(target_subdir)

            for filename in files:
                file_ext = filename.lower()
                source_file = os.path.join(root_dir, filename)

                if file_ext.endswith('.docx'):
                    base_name = os.path.splitext(filename)[0]
                    target_file = os.path.join(target_subdir, base_name + '.md')

                    result_text.insert(tk.END, f"Converting docx: {filename} -> {base_name}.md\n")
                    result_text.see(tk.END)
                    root.update()

                    if docx_to_markdown(source_file, target_file):
                        success_count += 1
                    else:
                        fail_count += 1
                        result_text.insert(tk.END, f"  Failed: {filename}\n")
                    total_files += 1

                elif file_ext.endswith('.pptx'):
                    base_name = os.path.splitext(filename)[0]
                    target_file = os.path.join(target_subdir, base_name + '.md')

                    result_text.insert(tk.END, f"Converting pptx: {filename} -> {base_name}.md\n")
                    result_text.see(tk.END)
                    root.update()

                    if pptx_to_markdown(source_file, target_file):
                        success_count += 1
                    else:
                        fail_count += 1
                        result_text.insert(tk.END, f"  Failed: {filename}\n")
                    total_files += 1

        result_text.insert(tk.END, f"\nProcessing complete!\n")
        result_text.insert(tk.END, f"Total files: {total_files}\n")
        result_text.insert(tk.END, f"Success: {success_count}\n")
        result_text.insert(tk.END, f"Failed: {fail_count}\n")
        root.update()

        status_label.config(text="Processing complete")
        messagebox.showinfo("Complete", f"Processing complete!\nTotal files: {total_files}\nSuccess: {success_count}\nFailed: {fail_count}")

    except Exception as e:
        result_text.insert(tk.END, f"\nError: {str(e)}\n")
        root.update()
        status_label.config(text="Processing failed")
        messagebox.showerror("Error", f"Processing failed: {str(e)}")
    finally:
        btn_process.config(state="normal")


def create_gui():
    global source_entry, target_entry, btn_process, status_label, result_text, root

    root = tk.Tk()
    root.title("Batch Document to Markdown Tool")
    root.geometry("800x600")

    SKY_BLUE = "#FAF9F6"
    LIGHT_BLUE = "#E8E4E1"
    DARK_BLUE = "#4A4A4A"
    WHITE = "#FFFFFF"

    root.configure(bg=SKY_BLUE)

    style = ttk.Style()
    style.configure("TFrame", background=SKY_BLUE)
    style.configure("TLabel", background=SKY_BLUE, foreground="#333333")
    style.configure("TEntry", fieldbackground=WHITE)
    style.configure("TButton", background=LIGHT_BLUE, foreground="#333333", font=("Arial", 12))

    frame = ttk.Frame(root, padding="20")
    frame.pack(fill=tk.BOTH, expand=True)
    frame.configure(style="TFrame")

    title_label = ttk.Label(frame, text="Batch Document to Markdown Tool", font=("Arial", 16, "bold"))
    title_label.pack(anchor=tk.CENTER, pady=(0, 20))

    # Source Directory
    source_frame = ttk.Frame(frame)
    source_frame.pack(fill=tk.X, pady=(0, 10))

    ttk.Label(source_frame, text="Source Directory:", font=("Arial", 11)).pack(anchor=tk.W, pady=(0, 5))

    source_entry_frame = ttk.Frame(source_frame)
    source_entry_frame.pack(fill=tk.X)

    source_entry = ttk.Entry(source_entry_frame, font=("Arial", 11))
    source_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

    source_btn = ttk.Button(source_entry_frame, text="Browse...", command=select_source_dir)
    source_btn.pack(side=tk.RIGHT)

    # Target Directory
    target_frame = ttk.Frame(frame)
    target_frame.pack(fill=tk.X, pady=(0, 20))

    ttk.Label(target_frame, text="Target Directory:", font=("Arial", 11)).pack(anchor=tk.W, pady=(0, 5))

    target_entry_frame = ttk.Frame(target_frame)
    target_entry_frame.pack(fill=tk.X)

    target_entry = ttk.Entry(target_entry_frame, font=("Arial", 11))
    target_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

    target_btn = ttk.Button(target_entry_frame, text="Browse...", command=select_target_dir)
    target_btn.pack(side=tk.RIGHT)

    # Description
    info_label = ttk.Label(frame, text="Description: Scans docx and pptx files in all subdirectories of the source directory and converts them to md format in the target directory",
                           font=("Arial", 9), foreground="#666666")
    info_label.pack(anchor=tk.W, pady=(0, 10))

    # Process Button
    btn_process = tk.Button(frame, font=("Arial", 12, "bold"),
                            bg=DARK_BLUE, fg=WHITE,
                            padx=20, pady=8, relief=tk.RAISED, bd=2,
                            text="Start Processing", command=process_files)
    btn_process.pack(fill=tk.X, pady=(0, 20))

    # Status Label
    status_label = ttk.Label(frame, text="Ready", font=("Arial", 10))
    status_label.pack(anchor=tk.W, pady=(0, 5))

    # Progress Label
    ttk.Label(frame, text="Processing Progress:", font=("Arial", 11)).pack(anchor=tk.W, pady=(0, 5))

    # Result Text Box
    result_text = tk.Text(frame, font=("Consolas", 10), height=15, bg=WHITE, fg="#333333")
    result_text.pack(fill=tk.BOTH, expand=True)

    scrollbar = ttk.Scrollbar(result_text, command=result_text.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    result_text.config(yscrollcommand=scrollbar.set)

    root.mainloop()


if __name__ == "__main__":
    create_gui()
