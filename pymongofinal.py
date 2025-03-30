import pymongo
from collections import Counter
import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd

# MongoDB connection
mongo_client = pymongo.MongoClient("mongodb://localhost:7000/")
mongo_db = mongo_client['groupwmongo']
employee_feedback_col = mongo_db['employee_feedback']
offboarding_review_col = mongo_db['offboarding_review']
employee_development_col = mongo_db['employee_personal_development']
reprimands_col = mongo_db['reprimands']

try:
    # Graph 1 key reprimand types 
    # Fetch the reprimands from the collection and count by REPRIMAND_TYPE
    reprimand_types = [reprimand['REPRIMAND_TYPE'] for reprimand in reprimands_col.find()]

    # Use Counter to get the frequency of each reprimand type
    #reprimand_types = [reprimand.get('REPRIMAND_TYPE', 'Unknown') for reprimand in reprimands_col.find()]
    reprimand_counts = Counter(reprimand_types)

    # Extract the reprimand types and counts
    labels = list(reprimand_counts.keys())
    values = list(reprimand_counts.values())

    # Plot the bar graph
    plt.figure(figsize=(10, 6))  
    plt.bar(labels, values, color='skyblue')
    plt.xlabel('Reprimand Type')
    plt.ylabel('Count')
    plt.title('Count of Reprimands by Type')
    plt.xticks(rotation=45)  # Rotate the x labels for better visibility
    plt.tight_layout()  # Adjust layout to fit labels
    plt.show()



    employee_development_list = list(employee_development_col.find())  # This will get all the documents

    #Convert the data to a DataFrame for easier analysis and visualization
    employee_data = []
    for employee in employee_development_list:
        for program in employee["DEVELOPMENT_PROGRAMS"]:
            for goal in employee["PERFORMANCE_GOALS"]:
                employee_data.append({
                    "EMPLOYEE_ID": employee.get("EMPLOYEE_ID"),
                    "PROGRAM_NAME": program.get("PROGRAM_NAME"),
                    "STATUS": program.get("STATUS"),
                    "SKILLS_ACQUIRED": ", ".join(program.get("SKILLS_ACQUIRED")),
                    "IMPACT": program.get("IMPACT"),
                    "GOAL": goal.get("GOAL"),
                    "TARGET_DATE": goal.get("TARGET_DATE"),
                    "PROGRESS":goal.get("PROGRESS"),
                    "COMMENTS": goal.get("COMMENTS"),
                    "AREAS_FOR_IMPROVEMENT": ", ".join(employee["AREAS_FOR_IMPROVEMENT"]),
                })

    # Check the employee_data before creating the DataFrame
    #print(f"Employee data: {employee_data[:5]}")  # Show first 5 rows of employee_data

    df = pd.DataFrame(employee_data)
    df.head()
    
    # Graph 2 Status Distribution of Development Programs
    status_counts = df["STATUS"].value_counts()
    plt.figure(figsize=(7, 7))
    plt.pie(status_counts, labels=status_counts.index, autopct='%1.1f%%', startangle=90, colors=['#66b3ff', '#99ff99', '#ffcc99'])
    plt.title('Development Program Status Distribution')
    plt.axis('equal')
    plt.show()

    #Graph 3 Status of employees with remaining progress
    #Fill NaN values with '0%' and convert 'PROGRESS' to integers
    df['PROGRESS'] = df['PROGRESS'].fillna('0%')  # Fill NaN values with '0%'
    df['PROGRESS'] = df['PROGRESS'].apply(lambda x: int(x.replace('%', '').strip()) if isinstance(x, str) and '%' in x else int(x))

    # Plot 1: Progress towards Performance Goals (Horizontal Bar Chart)
    plt.figure(figsize=(10, 6))

    # Switch GOAL to y-axis and PROGRESS to x-axis for a horizontal bar chart
    sns.barplot(x="PROGRESS", y="GOAL", data=df, hue="EMPLOYEE_ID", dodge=False, ci=None)

    plt.title('Progress Towards Performance Goals by Goal')
    plt.xlabel('Progress (%)')
    plt.ylabel('Performance Goal')
    plt.legend(title='Employee ID', bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.tight_layout()
    plt.show()



except Exception as e:
    # Handle general exceptions
    print(f"An error occurred: {e}")

finally:
# Ensure MongoDB connection is closed
    if 'mongo_client' in locals():
        mongo_client.close()  # Close the MongoDB connection
        print("MongoDB connection closed.")
