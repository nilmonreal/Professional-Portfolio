import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

class BudgetVsActualPlot:
    def __init__(self, file_path):
        self.file_path = file_path
        self.data = None
        self.data_melted = None

    def load_data(self): # Loads and cleans the data from the CSV file.
        try:
            # Load budget vs actual data
            self.data = pd.read_csv(self.file_path, header=0)
            # Strip any leading/trailing whitespace from column names
            self.data.columns = self.data.columns.str.strip()
            print("Data loaded successfully.")
        except FileNotFoundError:
            print(f"Error: The file at {self.file_path} was not found.")
        except pd.errors.EmptyDataError:
            print(f"Error: The file at {self.file_path} is empty.")
        except Exception as e:
            print(f"An error occurred: {e}")

    def process_data(self): #Calculates variance and reshapes the data for plotting.
        try:
            # Calculate variance and percentage variance
            self.data['Variance'] = self.data['Actual'] - self.data['Budget']
            self.data['% Variance'] = (self.data['Variance'] / self.data['Budget']) * 100
            # Reshape the data for seaborn plotting
            self.data_melted = pd.melt(self.data, id_vars=['Metric'], value_vars=['Budget', 'Actual'], var_name='Type', value_name='Amount')
            print("Data processed successfully.")
        except KeyError as e:
            print(f"Error: Missing expected column {e}")
        except Exception as e:
            print(f"An error occurred: {e}")    

    def plot_data(self): #Generates the bar plot using Seaborn.
        try:
            # Set the Seaborn style
            sns.set(style="whitegrid")
            # Create a bar plot with Seaborn
            plt.figure(figsize=(12, 6))
            sns.barplot(x='Metric', y='Amount', hue='Type', data=self.data_melted, palette=['#00ADB5', '#393E46'])
            # Customize the plot
            plt.xlabel('GL', fontweight='bold')
            plt.xticks(rotation=45, ha='right')
            plt.ylabel('Amount in $')
            plt.title('Budget vs Actual')
            plt.legend(title='Type')
            plt.tight_layout(pad=2.0)  # Adjust layout to add padding and margins
            plt.show()
        except Exception as e:
            print(f"An error occurred while plotting the data: {e}")


# Test the class and load_data method
plotter = BudgetVsActualPlot(r'C:\Users\nil.monreal\Desktop\Nil\EZ Notes\Python Scripts\Excel Files\P&L Actual vs Budget.csv')
plotter.load_data()
plotter.process_data()
plotter.plot_data()
