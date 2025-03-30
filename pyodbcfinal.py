import pyodbc 
import matplotlib.pyplot as plt 
import numpy as np 
 
# Database connection 
server = 'tcp:mcruebs04.isad.isadroot.ex.ac.uk'  
database = 'BEMM459_GroupW' 
username = 'GroupW'  
password = 'YhdF813+Kr' 
 
serverstring = 'DRIVER={ODBC Driver 18 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password+';TrustServerCertificate=yes;Encrypt=no;' 
 
try: 
    cnxn = pyodbc.connect(serverstring) 
    cursor = cnxn.cursor() 
    print("Connected to the database successfully!") 
 
    # Employee Skill Proficiency Distribution Querying and visualisation 
    cursor.execute(""" 
        SELECT s.skill_name, es.proficiency_level, COUNT(*) AS SkillCount 
        FROM [HROffboarding].[EMPLOYEE_SKILLS] es 
        JOIN [HROffboarding].[SKILL_COMPENDIUM] s ON es.skill_id = s.skill_id 
        GROUP BY s.skill_name, es.proficiency_level 
        ORDER BY s.skill_name, SkillCount DESC; 
    """) 
    results1 = cursor.fetchall() 
 
    skills = list(set([row.skill_name for row in results1])) 
    proficiency_levels = list(set([row.proficiency_level for row in results1])) 
    data = {} 
    for skill in skills: 
        data[skill] = {level: 0 for level in proficiency_levels} 
    for row in results1: 
        data[row.skill_name][row.proficiency_level] = row.SkillCount 
 
    width = 0.2 
    x = np.arange(len(skills)) 
    plt.figure(figsize=(12, 8)) 
    for i, level in enumerate(proficiency_levels): 
        values = [data[skill][level] for skill in skills] 
        plt.bar(x + i * width, values, width, label=level) 
 
    plt.xlabel('Skills') 
    plt.ylabel('Count') 
    plt.title('Employee Skill Proficiency Distribution') 
    plt.xticks(x + width * (len(proficiency_levels) - 1) / 2, skills, rotation=45, ha='right') 
    plt.legend() 
    plt.tight_layout() 
    plt.show() 
 
    # Employees with the Highest Non-Billable Hours 
    cursor.execute(""" 
        SELECT e.employeeID, e.firstName, e.lastName, SUM(p.nonBillableHours) AS TotalNonBillableHours 
        FROM [HROffboarding].[EMPLOYEE] e 
        JOIN [HROffboarding].[PAYROLL] p ON e.employeeID = p.employeeID 
        GROUP BY e.employeeID, e.firstName, e.lastName 
        ORDER BY TotalNonBillableHours DESC; 
    """) 
    results2 = cursor.fetchall() 
    employee_names = [f"{row.firstName} {row.lastName}" for row in results2] 
    non_billable_hours = [row.TotalNonBillableHours for row in results2] 
 
    plt.figure(figsize=(10, 6)) 
    plt.bar(employee_names[:10], non_billable_hours[:10])  # Top 10 
    plt.xlabel('Employee') 
    plt.ylabel('Total Non-Billable Hours') 
    plt.title('Employees with High Non-Billable Hours') 
    plt.xticks(rotation=45, ha='right') 
    plt.tight_layout() 
    plt.show() 
 
    # Average Salary by Department 
    cursor.execute(""" 
        SELECT d.departmentName, AVG(p.salary) AS AverageSalary 
        FROM [HROffboarding].[EMPLOYEE] e 
        JOIN [HROffboarding].[PAYROLL] p ON e.employeeID = p.employeeID 
        JOIN [HROffboarding].[DEPARTMENT] d ON e.departmentID = d.departmentID 
        GROUP BY d.departmentName 
        ORDER BY AverageSalary DESC; 
    """) 
    results4 = cursor.fetchall() 
    departments8 = [row.departmentName for row in results4] 
    average_salaries = [row.AverageSalary for row in results4] 
 
    plt.figure(figsize=(10, 6)) 
    plt.bar(departments8, average_salaries) 
    plt.xlabel('Department') 
    plt.ylabel('Average Salary') 
    plt.title('Average Salary by Department') 
    plt.xticks(rotation=45, ha='right') 
    plt.tight_layout() 
    plt.show() 
 
    # Employee Performance Ratings Distribution 
    cursor.execute(""" 
        SELECT performanceRating 
        FROM [HROffboarding].[PERFORMANCE_DATA]; 
    """) 
    results5 = cursor.fetchall() 
    ratings12 = [row.performanceRating for row in results5] 
 
    plt.figure(figsize=(8, 6)) 
    plt.hist(ratings12, bins=np.arange(min(ratings12), max(ratings12) + 1.5) - 0.5, edgecolor='black') 
    plt.xlabel('Performance Rating') 
    plt.ylabel('Frequency') 
    plt.title('Performance Ratings Distribution') 
    plt.tight_layout() 
    plt.show() 
 
    # Top 5 Skills by Frequency 
    cursor.execute(""" 
        SELECT TOP 5 s.skill_name, COUNT(*) AS SkillCount 
        FROM [HROffboarding].[EMPLOYEE_SKILLS] es 
        JOIN [HROffboarding].[SKILL_COMPENDIUM] s ON es.skill_id = s.skill_id 
        GROUP BY s.skill_name 
        ORDER BY SkillCount DESC; 
    """) 
    results6 = cursor.fetchall() 
    skills13 = [row.skill_name for row in results6] 
    skill_counts13 = [row.SkillCount for row in results6] 
 
    plt.figure(figsize=(8, 8)) 
    plt.pie(skill_counts13, labels=skills13, autopct='%1.1f%%', startangle=140) 
    plt.title('Top 5 Skills by Frequency') 
    plt.tight_layout() 
    plt.show() 
 
    # Employee Bill Rate vs. Revenue Generated 
    cursor.execute(""" 
        SELECT e.billRate, pd.revenueGenerated, pd.utilizationRate 
        FROM [HROffboarding].[EMPLOYEE] e 
        JOIN [HROffboarding].[PERFORMANCE_DATA] pd ON e.employeeID = pd.employeeID; 
    """) 
    results9 = cursor.fetchall() 
    bill_rates15 = [row.billRate for row in results9] 
    revenue15 = [row.revenueGenerated for row in results9] 
    utilization15 = [row.utilizationRate for row in results9] 
 
    plt.figure(figsize=(10, 6)) 
    plt.scatter(bill_rates15, revenue15, s=[u * 10 for u in utilization15], alpha=0.5) 
    plt.xlabel('Bill Rate') 
    plt.ylabel('Revenue Generated') 
    plt.title('Employee Bill Rate vs. Revenue Generated') 
    plt.colorbar(label='Utilization Rate') 
    plt.tight_layout() 
    plt.show() 
 
except pyodbc.Error as e: 
    print(f" Database error: {e}") 
 
finally: 
    if 'cnxn' in locals(): 
        cnxn.close() 
        print("Database connection closed.") 