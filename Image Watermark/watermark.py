import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from PIL import Image, ImageTk, ImageFont, ImageDraw
import os

class WatermarkApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Image Watermarker")
        self.root.geometry("800x600")

        # Image display area
        self.image_label = tk.Label(root, text="No image selected")
        self.image_label.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)

        # Button frame
        button_frame = tk.Frame(root)
        button_frame.pack(pady=10)

        # Buttons
        upload_btn = tk.Button(button_frame, text="Upload Image", command=self.upload_image)
        upload_btn.pack(side=tk.LEFT, padx=10)

        text_watermark_btn = tk.Button(button_frame, text="Add Text Watermark", command=self.add_text_watermark)
        text_watermark_btn.pack(side=tk.LEFT, padx=10)

        image_watermark_btn = tk.Button(button_frame, text="Add Image Watermark", command=self.add_image_watermark)
        image_watermark_btn.pack(side=tk.LEFT, padx=10)

        save_btn = tk.Button(button_frame, text="Save Image", command=self.save_image)
        save_btn.pack(side=tk.LEFT, padx=10)

        # Instance variables
        self.original_image = None
        self.displayed_image = None
        self.current_image = None

    def upload_image(self):
        """Open file dialog to upload an image"""
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.gif *.bmp"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            try:
                # Open the image
                self.original_image = Image.open(file_path)
                self.current_image = self.original_image.copy()
                
                # Resize image to fit the display
                self.displayed_image = self.resize_image(self.current_image)
                
                # Convert to PhotoImage for display
                photo = ImageTk.PhotoImage(self.displayed_image)
                
                # Update label
                self.image_label.configure(image=photo)
                self.image_label.image = photo
            except Exception as e:
                messagebox.showerror("Error", f"Could not open image: {str(e)}")

    def resize_image(self, image, max_width=600, max_height=400):
        """Resize image to fit display while maintaining aspect ratio"""
        image.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)
        return image

    def add_text_watermark(self):
        """Add text watermark to the image"""
        if self.current_image is None:
            messagebox.showwarning("Warning", "Please upload an image first")
            return

        # Ask for watermark text
        text = simpledialog.askstring("Input", "Enter watermark text:")
        if text:
            # Create a copy of the image to modify
            watermarked = self.current_image.copy()
            
            # Prepare drawing context
            draw = ImageDraw.Draw(watermarked)
            
            # Try to use a default font
            try:
                font = ImageFont.truetype("arial.ttf", 36)
            except IOError:
                # Fallback to default font
                font = ImageFont.load_default()
            
            # Get image dimensions
            width, height = watermarked.size
            
            # Calculate text position (bottom right with some padding)
            text_width, text_height = draw.textbbox((0,0), text, font=font)[2:4]
            position = (width - text_width - 20, height - text_height - 20)
            
            # Draw text with semi-transparent background
            draw.rectangle([position[0]-10, position[1]-10, 
                            position[0]+text_width+10, position[1]+text_height+10], 
                           fill=(255,255,255,128))
            
            # Draw text
            draw.text(position, text, font=font, fill=(0,0,0,200))
            
            # Update current image and display
            self.current_image = watermarked
            self.displayed_image = self.resize_image(watermarked.copy())
            photo = ImageTk.PhotoImage(self.displayed_image)
            self.image_label.configure(image=photo)
            self.image_label.image = photo

    def add_image_watermark(self):
        """Add an image watermark"""
        if self.current_image is None:
            messagebox.showwarning("Warning", "Please upload an image first")
            return

        # Open watermark image
        watermark_path = filedialog.askopenfilename(
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.gif *.bmp"),
                ("All files", "*.*")
            ]
        )
        
        if watermark_path:
            try:
                # Open watermark image
                watermark = Image.open(watermark_path)
                
                # Create a copy of the original image
                watermarked = self.current_image.copy()
                
                # Resize watermark (make it smaller)
                watermark_width = int(watermarked.width * 0.2)  # 20% of image width
                watermark_aspect = watermark.width / watermark.height
                watermark_height = int(watermark_width / watermark_aspect)
                watermark = watermark.resize((watermark_width, watermark_height), Image.Resampling.LANCZOS)
                
                # Make watermark semi-transparent
                watermark = watermark.convert("RGBA")
                data = watermark.getdata()
                new_data = []
                for item in data:
                    # Reduce alpha channel to make semi-transparent
                    new_data.append((item[0], item[1], item[2], 100))
                watermark.putdata(new_data)
                
                # Calculate position (bottom right corner)
                position = (
                    watermarked.width - watermark.width - 20, 
                    watermarked.height - watermark.height - 20
                )
                
                # Paste watermark
                watermarked.paste(watermark, position, watermark)
                
                # Update current image and display
                self.current_image = watermarked
                self.displayed_image = self.resize_image(watermarked.copy())
                photo = ImageTk.PhotoImage(self.displayed_image)
                self.image_label.configure(image=photo)
                self.image_label.image = photo
            
            except Exception as e:
                messagebox.showerror("Error", f"Could not add watermark: {str(e)}")

    def save_image(self):
        """Save the watermarked image"""
        if self.current_image is None:
            messagebox.showwarning("Warning", "No image to save")
            return

        # Open save dialog
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[
                ("PNG files", "*.png"),
                ("JPEG files", "*.jpg"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            try:
                # Save the full-resolution original image with watermark
                self.current_image.save(file_path)
                messagebox.showinfo("Success", f"Image saved to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not save image: {str(e)}")

def main():
    root = tk.Tk()
    app = WatermarkApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()