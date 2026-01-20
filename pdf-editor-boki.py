import customtkinter as ctk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import os
from pathlib import Path

class PDFSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Splitter & Auto-Renamer")
        self.root.geometry("900x700")
        
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        self.pdf_path = None
        self.pdf_doc = None
        self.output_dir = None
        self.rect_start = None
        self.rect_end = None
        self.current_rect = None
        self.canvas_image = None
        self.page_image = None
        self.scale_factor = 1.0
        
        self.setup_ui()
    
    def setup_ui(self):
        # Top control panel
        control_frame = ctk.CTkFrame(self.root)
        control_frame.pack(pady=10, padx=10, fill="x")
        
        self.load_btn = ctk.CTkButton(control_frame, text="Load PDF", command=self.load_pdf)
        self.load_btn.pack(side="left", padx=5)
        
        self.output_btn = ctk.CTkButton(control_frame, text="Choose Output Folder", 
                                        command=self.choose_output, state="disabled")
        self.output_btn.pack(side="left", padx=5)
        
        self.process_btn = ctk.CTkButton(control_frame, text="Process & Rename", 
                                         command=self.process_pdf, state="disabled")
        self.process_btn.pack(side="left", padx=5)
        
        self.clear_btn = ctk.CTkButton(control_frame, text="Clear Rectangle", 
                                       command=self.clear_rectangle, state="disabled")
        self.clear_btn.pack(side="left", padx=5)
        
        # Info label
        self.info_label = ctk.CTkLabel(self.root, text="Load a PDF to begin", 
                                       font=("Arial", 12))
        self.info_label.pack(pady=5)
        
        # Canvas frame
        canvas_frame = ctk.CTkFrame(self.root)
        canvas_frame.pack(pady=10, padx=10, fill="both", expand=True)
        
        # Canvas with scrollbars
        self.canvas = ctk.CTkCanvas(canvas_frame, bg="#2b2b2b", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        
        # Bind mouse events
        self.canvas.bind("<ButtonPress-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)
        
        # Status label
        self.status_label = ctk.CTkLabel(self.root, text="Ready", font=("Arial", 10))
        self.status_label.pack(pady=5)
    
    def load_pdf(self):
        filepath = filedialog.askopenfilename(
            title="Select PDF",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if not filepath:
            return
        
        try:
            self.pdf_path = filepath
            self.pdf_doc = fitz.open(filepath)
            
            # Display first page
            self.display_first_page()
            
            self.info_label.configure(
                text=f"Loaded: {Path(filepath).name} ({len(self.pdf_doc)} pages)\n"
                     "Draw a rectangle around the text to use for naming"
            )
            
            self.output_btn.configure(state="normal")
            self.clear_btn.configure(state="normal")
            self.status_label.configure(text="Draw rectangle on the text area")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load PDF: {str(e)}")
    
    def display_first_page(self):
        if not self.pdf_doc:
            return
        
        # Render first page
        page = self.pdf_doc[0]
        
        # Calculate scale to fit canvas
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        
        if canvas_width < 100:  # Canvas not ready yet
            canvas_width = 800
            canvas_height = 600
        
        # Get page dimensions
        mat = fitz.Matrix(2.6, 2.6)  # 2.6x zoom for 30% more
        pix = page.get_pixmap(matrix=mat)
        
        # Convert to PIL Image
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # Scale to fit canvas
        img_width, img_height = img.size
        scale_w = (canvas_width - 40) / img_width
        scale_h = (canvas_height - 40) / img_height
        # Always zoom in at least 2.5x, but do not shrink below 1.0
        self.scale_factor = max(0.55, min(scale_w, scale_h))
        
        new_width = int(img_width * self.scale_factor)
        new_height = int(img_height * self.scale_factor)
        
        img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
        
        self.page_image = img
        self.canvas_image = ImageTk.PhotoImage(img)
        
        # Clear canvas and display image
        self.canvas.delete("all")
        self.canvas.create_image(20, 20, anchor="nw", image=self.canvas_image)
        
        # Store image offset for coordinate calculations
        self.img_offset_x = 20
        self.img_offset_y = 20
    
    def on_mouse_down(self, event):
        if not self.page_image:
            return
        self.rect_start = (event.x, event.y)
        if self.current_rect:
            self.canvas.delete(self.current_rect)
    
    def on_mouse_drag(self, event):
        if not self.rect_start:
            return
        
        if self.current_rect:
            self.canvas.delete(self.current_rect)
        
        x1, y1 = self.rect_start
        x2, y2 = event.x, event.y
        
        self.current_rect = self.canvas.create_rectangle(
            x1, y1, x2, y2, outline="#00ff00", width=2
        )
    
    def on_mouse_up(self, event):
        if not self.rect_start:
            return
        
        self.rect_end = (event.x, event.y)
        
        # Enable process button if output dir is selected
        if self.output_dir:
            self.process_btn.configure(state="normal")
    
    def clear_rectangle(self):
        if self.current_rect:
            self.canvas.delete(self.current_rect)
            self.current_rect = None
            self.rect_start = None
            self.rect_end = None
            self.process_btn.configure(state="disabled")
    
    def choose_output(self):
        dirpath = filedialog.askdirectory(title="Select Output Folder")
        
        if dirpath:
            self.output_dir = dirpath
            self.status_label.configure(text=f"Output: {dirpath}")
            
            if self.rect_start and self.rect_end:
                self.process_btn.configure(state="normal")
    
    def extract_text_from_rect(self, page, rect_coords):
        """Extract text from specified rectangle coordinates"""
        try:
            # Convert canvas coordinates to PDF coordinates
            x1 = (rect_coords[0] - self.img_offset_x) / self.scale_factor / 2
            y1 = (rect_coords[1] - self.img_offset_y) / self.scale_factor / 2
            x2 = (rect_coords[2] - self.img_offset_x) / self.scale_factor / 2
            y2 = (rect_coords[3] - self.img_offset_y) / self.scale_factor / 2
            
            # Create rectangle
            rect = fitz.Rect(x1, y1, x2, y2)
            
            # Extract text
            text = page.get_text("text", clip=rect).strip()
            
            # Clean filename
            text = "".join(c for c in text if c.isalnum() or c in (' ', '-', '_')).strip()
            text = text.replace(' ', '_')
            
            return text if text else None
            
        except Exception as e:
            print(f"Error extracting text: {e}")
            return None
    
    def process_pdf(self):
        if not self.pdf_doc or not self.output_dir or not self.rect_start or not self.rect_end:
            messagebox.showwarning("Warning", "Please complete all steps first")
            return
        
        try:
            # Get rectangle coordinates
            x1 = min(self.rect_start[0], self.rect_end[0])
            y1 = min(self.rect_start[1], self.rect_end[1])
            x2 = max(self.rect_start[0], self.rect_end[0])
            y2 = max(self.rect_start[1], self.rect_end[1])
            rect_coords = (x1, y1, x2, y2)
            
            # Track used names for duplicates
            name_count = {}
            
            total_pages = len(self.pdf_doc)
            self.status_label.configure(text="Processing...")
            self.root.update()
            
            for page_num in range(total_pages):
                page = self.pdf_doc[page_num]
                
                # Extract text from rectangle
                text = self.extract_text_from_rect(page, rect_coords)
                
                # Determine filename
                if text:
                    base_name = text
                    if base_name in name_count:
                        name_count[base_name] += 1
                        filename = f"{base_name}_{name_count[base_name]}.pdf"
                    else:
                        name_count[base_name] = 1
                        filename = f"{base_name}.pdf"
                else:
                    filename = f"page_{page_num + 1}.pdf"
                
                # Create single-page PDF
                output_pdf = fitz.open()
                output_pdf.insert_pdf(self.pdf_doc, from_page=page_num, to_page=page_num)
                
                # Save
                output_path = os.path.join(self.output_dir, filename)
                output_pdf.save(output_path)
                output_pdf.close()
                
                # Update progress
                self.status_label.configure(
                    text=f"Processing... {page_num + 1}/{total_pages}"
                )
                self.root.update()
            
            self.status_label.configure(text=f"Complete! Created {total_pages} files")
            messagebox.showinfo("Success", 
                              f"Successfully created {total_pages} PDF files in:\n{self.output_dir}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Processing failed: {str(e)}")
            self.status_label.configure(text="Error occurred")

def main():
    root = ctk.CTk()
    app = PDFSplitterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()