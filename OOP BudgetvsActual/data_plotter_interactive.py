import pandas as pd
import plotly.express as px

class BudgetVsActualPlot:
    def __init__(self, file_path):
        self.file_path = file_path
        self.data = None
        self.data_melted = None

    def load_data(self):
        try:
            # Load budget vs actual data
            self.data = pd.read_csv(self.file_path, header=0)
            # Strip any leading/trailing whitespace from column names
            self.data.columns = self.data.columns.str.strip()
            print("Data loaded successfully.")
            print(self.data.head())
        except FileNotFoundError:
            print(f"Error: The file at {self.file_path} was not found.")
        except pd.errors.EmptyDataError:
            print(f"Error: The file at {self.file_path} is empty.")
        except Exception as e:
            print(f"An error occurred: {e}")

    def process_data(self):
        try:
            # Calculate variance and percentage variance
            self.data['Variance'] = self.data['Actual'] - self.data['Budget']
            self.data['% Variance'] = (self.data['Variance'] / self.data['Budget']) * 100
            # Reshape the data for Plotly plotting
            self.data_melted = pd.melt(self.data, id_vars=['Metric'], value_vars=['Budget', 'Actual'], var_name='Type', value_name='Amount')
            print("Data processed successfully.")
            print(self.data_melted.head())
        except KeyError as e:
            print(f"Error: Missing expected column {e}")
        except Exception as e:
            print(f"An error occurred: {e}")

    def plot_data(self):
        try:
            # Create an interactive bar plot with Plotly
            fig = px.bar(self.data_melted, x='Metric', y='Amount', color='Type', barmode='group',
                         labels={'Metric': 'GL', 'Amount': 'Amount in $'}, 
                         title='Budget vs Actual',
                         color_discrete_map={'Budget': '#00ADB5', 'Actual': '#393E46'})

            # Customize the layout
            fig.update_layout(
                xaxis_title='GL',
                yaxis_title='Amount in $',
                legend_title_text='Type',
                xaxis_tickangle=-45
            )

            fig.show()
        except Exception as e:
            print(f"An error occurred while plotting the data: {e}")

# Test the class and methods
plotter = BudgetVsActualPlot(r'C:\Users\nil.monreal\Desktop\Nil\EZ Notes\Python Scripts\Excel Files\P&L Actual vs Budget.csv')
plotter.load_data()
plotter.process_data()
plotter.plot_data()
