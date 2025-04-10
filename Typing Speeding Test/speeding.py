import tkinter as tk
from tkinter import messagebox
import random
import time

class TypingSpeedTest:
    def __init__(self, root):
        self.root = root
        self.root.title("Typing Speed Test")
        self.root.geometry("800x600")

        # Sample paragraphs for typing test
        self.paragraphs = [
            "The quick brown fox jumps over the lazy dog. This classic pangram contains every letter of the alphabet.",
            "Programming is not about what you know; it's about what you can figure out and learn along the way.",
            "Success is not final, failure is not fatal: it is the courage to continue that counts.",
            "In the world of technology, innovation is the key to solving complex problems and creating new opportunities.",
            "Learning never exhausts the mind. The more that you read, the more things you will know.",
            "Creativity is intelligence having fun. The best way to predict the future is to create it.",
            "Perseverance is not a long race; it is many short races one after the other.",
            "The only way to do great work is to love what you do and believe in yourself.",
        ]

        # Game state variables
        self.current_paragraph = ""
        self.start_time = 0
        self.is_test_active = False

        # Create UI components
        self.create_widgets()

    def create_widgets(self):
        """Create and layout all UI widgets."""
        # Paragraph display
        self.paragraph_label = tk.Label(
            self.root, 
            text="Click 'Start Test' to begin", 
            font=("Courier", 14), 
            wraplength=700, 
            justify=tk.LEFT
        )
        self.paragraph_label.pack(pady=20)

        # Input text area
        self.input_text = tk.Text(
            self.root, 
            height=8, 
            width=80, 
            font=("Courier", 12), 
            state=tk.DISABLED
        )
        self.input_text.pack(pady=10)

        # Buttons frame
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)

        # Start Test Button
        self.start_button = tk.Button(
            button_frame, 
            text="Start Test", 
            command=self.start_test,
            font=("Arial", 12)
        )
        self.start_button.pack(side=tk.LEFT, padx=10)

        # Stop Test Button
        self.stop_button = tk.Button(
            button_frame, 
            text="Stop Test", 
            command=self.stop_test,
            state=tk.DISABLED,
            font=("Arial", 12)
        )
        self.stop_button.pack(side=tk.LEFT, padx=10)

        # Results Frame
        results_frame = tk.Frame(self.root)
        results_frame.pack(pady=10)

        # Results Labels
        self.wpm_label = tk.Label(results_frame, text="WPM: -", font=("Arial", 12))
        self.wpm_label.pack(side=tk.LEFT, padx=10)

        self.accuracy_label = tk.Label(results_frame, text="Accuracy: -", font=("Arial", 12))
        self.accuracy_label.pack(side=tk.LEFT, padx=10)

        self.time_label = tk.Label(results_frame, text="Time: -", font=("Arial", 12))
        self.time_label.pack(side=tk.LEFT, padx=10)

    def start_test(self):
        """Initialize and start the typing test."""
        # Select a random paragraph
        self.current_paragraph = random.choice(self.paragraphs)
        
        # Display paragraph
        self.paragraph_label.config(text=self.current_paragraph)
        
        # Enable text input
        self.input_text.config(state=tk.NORMAL)
        self.input_text.delete('1.0', tk.END)
        self.input_text.focus()
        
        # Manage button states
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        
        # Reset labels
        self.wpm_label.config(text="WPM: -")
        self.accuracy_label.config(text="Accuracy: -")
        self.time_label.config(text="Time: -")
        
        # Start timer
        self.start_time = time.time()
        self.is_test_active = True
        
        # Bind key release event to calculate progress
        self.input_text.bind('<KeyRelease>', self.calculate_progress)

    def stop_test(self):
        """Stop the typing test and calculate results."""
        if not self.is_test_active:
            return
        
        # Unbind key release event
        self.input_text.unbind('<KeyRelease>')
        
        # Calculate test results
        typed_text = self.input_text.get('1.0', tk.END).strip()
        end_time = time.time()
        
        # Disable text input
        self.input_text.config(state=tk.DISABLED)
        
        # Manage button states
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        
        # Calculate metrics
        self.calculate_metrics(typed_text, end_time)
        
        # Reset test state
        self.is_test_active = False

    def calculate_progress(self, event):
        """Real-time progress calculation during typing."""
        if not self.is_test_active:
            return
        
        # Calculate time elapsed
        current_time = time.time()
        elapsed_time = current_time - self.start_time
        
        # Get typed text
        typed_text = self.input_text.get('1.0', tk.END).strip()
        
        # Estimate WPM and time
        if elapsed_time > 0:
            words_typed = len(typed_text.split())
            estimated_wpm = int((words_typed / elapsed_time) * 60)
            self.time_label.config(text=f"Time: {elapsed_time:.1f} sec")
            self.wpm_label.config(text=f"WPM: {estimated_wpm}")

    def calculate_metrics(self, typed_text, end_time):
        """Calculate detailed typing test metrics."""
        # Calculate time
        elapsed_time = end_time - self.start_time
        minutes = elapsed_time / 60
        
        # Calculate Words Per Minute (WPM)
        words_typed = len(typed_text.split())
        wpm = int(words_typed / minutes) if minutes > 0 else 0
        
        # Calculate Accuracy
        correct_chars = sum(
            1 for a, b in zip(typed_text, self.current_paragraph) if a == b
        )
        total_chars = len(self.current_paragraph)
        accuracy = (correct_chars / total_chars) * 100 if total_chars > 0 else 0
        
        # Update labels
        self.wpm_label.config(text=f"WPM: {wpm}")
        self.accuracy_label.config(text=f"Accuracy: {accuracy:.1f}%")
        self.time_label.config(text=f"Time: {elapsed_time:.1f} sec")
        
        # Show detailed results
        messagebox.showinfo(
            "Typing Test Results", 
            f"Words Per Minute (WPM): {wpm}\n"
            f"Accuracy: {accuracy:.1f}%\n"
            f"Total Time: {elapsed_time:.1f} seconds"
        )

def main():
    root = tk.Tk()
    app = TypingSpeedTest(root)
    root.mainloop()

if __name__ == "__main__":
    main()