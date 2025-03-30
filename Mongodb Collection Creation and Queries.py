#Connectivity
# Install pymongo package
import subprocess
import sys

def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])
    
install("pymongo")
import pymongo

mongoclient = pymongo.MongoClient("mongodb://localhost:7000/")
#You can also check databases that presently exist using a loop
dblist = mongoclient.list_database_names()
for x in dblist:
    print(x)    


#Defining functions 
def check_DatabaseExists(argDBName):
    local_dblist = mongoclient.list_database_names()
    if argDBName in local_dblist:
        print("The database ", argDBName, " exists.")
    else:
        print("The database ", argDBName, " does not exist.")

def check_CollectionExists(argDBName, argCollName, local_mydb):
    local_collist = local_mydb.list_collection_names()
    if argCollName in local_collist:
        print("The collection ",  argCollName, "exists in database ", argDBName)
    else:
        print("The collection ", argCollName, " does not exist in database ", argDBName)


#Create a new database       
mydb = mongoclient["groupwmongo"]
print(type(mydb))

#Check if database exists by calling function check_DatabaseExists with name of database as the arguement
check_DatabaseExists("groupwmongo")
print()


#Collections
#Return a list of all collections in your database:
collections = mydb.list_collection_names()
print(mydb.list_collection_names())

# Drop the collections if they exist to ensure  no repetition
if "employee_feedback" in collections:
    mydb["employee_feedback"].drop()  # Drop directly using mydb["collection_name"]
    print("employee_feedback collection dropped.")

if "offboarding_review" in collections:
    mydb["offboarding_review"].drop()  # Drop directly using mydb["collection_name"]
    print("offboarding_review collection dropped.")

if "employee_development" in collections:
    mydb["employee_development"].drop()  # Drop directly using mydb["collection_name"]
    print("employee_development collection dropped.")

if "reprimands" in collections:
    mydb["reprimands"].drop()  # Drop directly using mydb["collection_name"]
    print("reprimands collection dropped.")

#Create a new collection called "employee_feedback"
employee_feedback_col = mydb["employee_feedback"]
print(type(employee_feedback_col))

check_CollectionExists("groupwmongo", "employee_feedback", mydb)

print()
offboarding_review_col = mydb["offboarding_review"]
print(type(offboarding_review_col))
employee_development_col = mydb["employee_personal_development"]
print(type(employee_development_col))
reprimands_col = mydb["reprimands"]
print(type(reprimands_col))

# Inserting data
feedback_list = [
  { "EMPLOYEE_ID": 1, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-01", "FEEDBACK_TEXT": "This employee consistently delivers high-quality results and shows excellent leadership in the team." },
    { "FEEDBACK_DATE": "2022-02-01", "FEEDBACK_TEXT": "A dependable leader who fosters collaboration, but there’s room for further improvement in conflict resolution." },
    { "FEEDBACK_DATE": "2024-02-01", "FEEDBACK_TEXT": "Outstanding leadership, with strategic thinking that contributes significantly to company goals." },
    { "FEEDBACK_DATE": "2025-02-01", "FEEDBACK_TEXT": "Keeps driving success through motivational leadership and high-quality decision-making." }
  ] },

  { "EMPLOYEE_ID": 2, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-02", "FEEDBACK_TEXT": "This employee excels in data analysis and provides valuable insights, but communication skills need improvement." },
    { "FEEDBACK_DATE": "2022-02-02", "FEEDBACK_TEXT": "Has significantly improved in leading projects and communicating with diverse teams." },
    { "FEEDBACK_DATE": "2024-02-02", "FEEDBACK_TEXT": "Now an effective leader with the ability to present data in a digestible format for all stakeholders." }
  ] },

  { "EMPLOYEE_ID": 3, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-03", "FEEDBACK_TEXT": "This employee has strong potential but sometimes struggles with decision-making under pressure." },
    { "FEEDBACK_DATE": "2022-02-03", "FEEDBACK_TEXT": "Decision-making is more confident, but project management skills can still improve." },
    { "FEEDBACK_DATE": "2024-02-03", "FEEDBACK_TEXT": "Continues to grow into a solid manager with a clearer strategic approach." }
  ] },

  { "EMPLOYEE_ID": 4, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-04", "FEEDBACK_TEXT": "This employee is a great HR resource, though sometimes struggles with implementing new tools." },
    { "FEEDBACK_DATE": "2022-02-04", "FEEDBACK_TEXT": "Performance is strong, but continues to need work on team-building strategies." }
  ] },

  { "EMPLOYEE_ID": 5, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-05", "FEEDBACK_TEXT": "This employee is adaptable but needs to be more proactive in taking on leadership roles." },
    { "FEEDBACK_DATE": "2022-02-05", "FEEDBACK_TEXT": "Great growth in leadership, now frequently managing key projects." }
  ] },

  { "EMPLOYEE_ID": 6, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-06", "FEEDBACK_TEXT": "This employee's time management skills need improvement, as task delegation is often inconsistent." },
    { "FEEDBACK_DATE": "2022-02-06", "FEEDBACK_TEXT": "Delegation skills have improved, but balancing workload and prioritization are still key challenges." }
  ] },

  { "EMPLOYEE_ID": 7, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-07", "FEEDBACK_TEXT": "This employee demonstrates strong technical skills, but struggles with taking leadership roles." },
    { "FEEDBACK_DATE": "2022-02-07", "FEEDBACK_TEXT": "Demonstrates solid leadership in projects, but could improve in conflict management." }
  ] },

  { "EMPLOYEE_ID": 8, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-08", "FEEDBACK_TEXT": "This employee's technical skills are great, but communication with the team needs improvement." },
    { "FEEDBACK_DATE": "2022-02-08", "FEEDBACK_TEXT": "This employee is now a key influencer in the team, leading presentations and driving important decisions." }
  ] },

  { "EMPLOYEE_ID": 9, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-09", "FEEDBACK_TEXT": "This employee is a strong contributor, but struggles with managing multiple projects at once." },
    { "FEEDBACK_DATE": "2022-02-09", "FEEDBACK_TEXT": "Now handles multiple projects with ease, though attention to detail can still be improved." }
  ] },

  { "EMPLOYEE_ID": 10, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-10", "FEEDBACK_TEXT": "This employee has strong technical skills but struggles with collaborating across teams." },
    { "FEEDBACK_DATE": "2022-02-10", "FEEDBACK_TEXT": "This employee has become an outstanding team player, collaborating effectively and leading cross-functional initiatives." }
  ] },

  { "EMPLOYEE_ID": 11, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-11", "FEEDBACK_TEXT": "Employee often misses deadlines and needs to improve in prioritization." },
    { "FEEDBACK_DATE": "2022-02-11", "FEEDBACK_TEXT": "Has improved in task management and is now meeting deadlines more consistently." }
  ] },

  { "EMPLOYEE_ID": 12, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-12", "FEEDBACK_TEXT": "This employee has strong organizational skills but struggles with delegating tasks effectively." },
    { "FEEDBACK_DATE": "2023-02-12", "FEEDBACK_TEXT": "Has improved in delegating tasks and managing team workloads." }
  ] },

  { "EMPLOYEE_ID": 13, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-13", "FEEDBACK_TEXT": "Employee needs to improve in giving constructive feedback to peers." },
    { "FEEDBACK_DATE": "2022-02-13", "FEEDBACK_TEXT": "Now provides more constructive feedback and contributes to team discussions." }
  ] },

  { "EMPLOYEE_ID": 14, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-14", "FEEDBACK_TEXT": "This employee has a positive attitude but struggles with meeting performance targets." },
    { "FEEDBACK_DATE": "2023-02-14", "FEEDBACK_TEXT": "Performance has improved significantly, and this employee consistently meets targets now." }
  ] },

  { "EMPLOYEE_ID": 15, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-15", "FEEDBACK_TEXT": "Employee has strong problem-solving skills but often works independently instead of collaborating with the team." },
    { "FEEDBACK_DATE": "2023-02-15", "FEEDBACK_TEXT": "Has shown great improvement in collaboration and team involvement." }
  ] },

  { "EMPLOYEE_ID": 16, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-16", "FEEDBACK_TEXT": "Employee demonstrates leadership potential but needs to work on public speaking and presentation skills." },
    { "FEEDBACK_DATE": "2022-02-16", "FEEDBACK_TEXT": "Leadership skills have improved, and this employee now delivers presentations confidently." }
  ] },

  { "EMPLOYEE_ID": 17, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-17", "FEEDBACK_TEXT": "This employee has a strong grasp of technical concepts but should work on mentoring junior team members." },
    { "FEEDBACK_DATE": "2023-02-17", "FEEDBACK_TEXT": "Now mentors junior employees and shares technical knowledge effectively." }
  ] },

  { "EMPLOYEE_ID": 18, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-18", "FEEDBACK_TEXT": "Employee excels in individual tasks but struggles with team dynamics and communication." },
    { "FEEDBACK_DATE": "2022-02-18", "FEEDBACK_TEXT": "Great improvement in team communication and contributing to group discussions." }
  ] },

  { "EMPLOYEE_ID": 19, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-19", "FEEDBACK_TEXT": "This employee often struggles with time management and prioritizing tasks effectively." },
    { "FEEDBACK_DATE": "2022-02-19", "FEEDBACK_TEXT": "Improved time management and now balances multiple tasks more effectively." }
  ] },

  { "EMPLOYEE_ID": 20, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-20", "FEEDBACK_TEXT": "Employee is effective in task execution but can improve in taking on leadership roles." },
    { "FEEDBACK_DATE": "2023-02-20", "FEEDBACK_TEXT": "Has shown strong leadership and takes initiative in new projects." }
  ] },

  { "EMPLOYEE_ID": 21, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-21", "FEEDBACK_TEXT": "Employee has excellent technical skills but lacks confidence in decision-making." },
    { "FEEDBACK_DATE": "2023-02-21", "FEEDBACK_TEXT": "Has developed confidence in decision-making and is now a trusted advisor within the team." }
  ] },

  { "EMPLOYEE_ID": 22, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-22", "FEEDBACK_TEXT": "This employee has a strong work ethic but struggles with adapting to change." },
    { "FEEDBACK_DATE": "2022-02-22", "FEEDBACK_TEXT": "Has improved in adapting to change and now leads new initiatives effectively." }
  ] },

  { "EMPLOYEE_ID": 23, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-23", "FEEDBACK_TEXT": "This employee is an effective team member but needs to work on delegation and empowering others." },
    { "FEEDBACK_DATE": "2022-02-23", "FEEDBACK_TEXT": "Has become a more effective team leader by delegating tasks and empowering others." }
  ] },

  { "EMPLOYEE_ID": 24, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-24", "FEEDBACK_TEXT": "Employee is skilled but often lacks follow-through on tasks and projects." },
    { "FEEDBACK_DATE": "2023-02-24", "FEEDBACK_TEXT": "Has shown significant improvement in follow-through and now consistently completes tasks on time." }
  ] },

  { "EMPLOYEE_ID": 25, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-25", "FEEDBACK_TEXT": "This employee is a reliable contributor but could benefit from improved time management skills." },
    { "FEEDBACK_DATE": "2022-02-25", "FEEDBACK_TEXT": "Has significantly improved time management skills and now meets deadlines consistently." }
  ] },

  { "EMPLOYEE_ID": 26, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-26", "FEEDBACK_TEXT": "Employee is an excellent contributor but needs to work on delegating tasks to others." },
    { "FEEDBACK_DATE": "2022-02-26", "FEEDBACK_TEXT": "Has shown great improvement in delegating tasks and empowering other team members." }
  ] },

  { "EMPLOYEE_ID": 27, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-27", "FEEDBACK_TEXT": "This employee shows strong potential but often struggles with prioritizing tasks effectively." },
    { "FEEDBACK_DATE": "2023-02-27", "FEEDBACK_TEXT": "Now prioritizes tasks well and manages multiple projects with ease." }
  ] },

  { "EMPLOYEE_ID": 28, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-28", "FEEDBACK_TEXT": "Employee is great at problem-solving but should work on communication skills with the team." },
    { "FEEDBACK_DATE": "2022-02-28", "FEEDBACK_TEXT": "Has significantly improved communication and is now an effective team player." }
  ] },

  { "EMPLOYEE_ID": 29, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-02-29", "FEEDBACK_TEXT": "Employee is a strong technical contributor but often needs guidance on leadership and team management." },
    { "FEEDBACK_DATE": "2023-02-29", "FEEDBACK_TEXT": "Now actively contributes to leadership, guiding teams and managing projects effectively." }
  ] },

  { "EMPLOYEE_ID": 30, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-03-01", "FEEDBACK_TEXT": "This employee is great with clients but needs to improve their problem-solving in internal team issues." },
    { "FEEDBACK_DATE": "2022-03-01", "FEEDBACK_TEXT": "Has become a key problem solver in internal team issues and maintains strong client relations." }
  ] },

  { "EMPLOYEE_ID": 31, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-03-02", "FEEDBACK_TEXT": "Employee shows solid performance but could benefit from improving time management skills." },
    { "FEEDBACK_DATE": "2023-03-02", "FEEDBACK_TEXT": "Now meets deadlines consistently and manages time effectively across multiple tasks." }
  ] },

  { "EMPLOYEE_ID": 32, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-03-03", "FEEDBACK_TEXT": "Employee needs to develop better leadership and communication skills for team projects." },
    { "FEEDBACK_DATE": "2023-03-03", "FEEDBACK_TEXT": "Has improved leadership and communication skills, now leading team projects efficiently." }
  ] },

  { "EMPLOYEE_ID": 33, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-03-04", "FEEDBACK_TEXT": "Employee excels in technical tasks but often struggles with coordinating with other teams." },
    { "FEEDBACK_DATE": "2022-03-04", "FEEDBACK_TEXT": "Improved coordination with cross-functional teams and contributes actively to collaborative efforts." }
  ] },

  { "EMPLOYEE_ID": 34, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-03-05", "FEEDBACK_TEXT": "Employee has potential but needs to work on interpersonal relationships within the team." },
    { "FEEDBACK_DATE": "2023-03-05", "FEEDBACK_TEXT": "Has made notable improvements in team relationships and works effectively with peers." }
  ] },

  { "EMPLOYEE_ID": 35, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-03-06", "FEEDBACK_TEXT": "This employee demonstrates strong technical knowledge but needs to improve delegation and leadership skills." },
    { "FEEDBACK_DATE": "2022-03-06", "FEEDBACK_TEXT": "Has grown into a leadership role, improving delegation and managing team projects effectively." }
  ] },

  { "EMPLOYEE_ID": 36, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-03-07", "FEEDBACK_TEXT": "Employee often takes on too many tasks, which affects overall efficiency and task completion." },
    { "FEEDBACK_DATE": "2022-03-07", "FEEDBACK_TEXT": "Has learned to manage workload effectively and delegates tasks when necessary." }
  ] },

  { "EMPLOYEE_ID": 37, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-03-08", "FEEDBACK_TEXT": "Employee is dependable but struggles with offering creative solutions to problems." },
    { "FEEDBACK_DATE": "2022-03-08", "FEEDBACK_TEXT": "Has improved in offering innovative solutions and contributes to creative brainstorming sessions." }
  ] },

  { "EMPLOYEE_ID": 38, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-03-09", "FEEDBACK_TEXT": "This employee's work is solid but lacks engagement in team collaboration." },
    { "FEEDBACK_DATE": "2022-03-09", "FEEDBACK_TEXT": "Now actively engages in team collaboration and leads initiatives within the group." }
  ] },

  { "EMPLOYEE_ID": 39, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-03-10", "FEEDBACK_TEXT": "Employee is great at taking ownership but could improve in working collaboratively with cross-functional teams." },
    { "FEEDBACK_DATE": "2023-03-10", "FEEDBACK_TEXT": "Has significantly improved in cross-functional collaboration and now contributes to team success." }
  ] },

  { "EMPLOYEE_ID": 40, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-03-11", "FEEDBACK_TEXT": "This employee performs well under pressure but could work on developing a more strategic approach to projects." },
    { "FEEDBACK_DATE": "2023-03-11", "FEEDBACK_TEXT": "Has developed a strategic mindset and is able to balance short-term and long-term objectives." }
  ] },

  { "EMPLOYEE_ID": 41, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-03-12", "FEEDBACK_TEXT": "Employee's technical skills are excellent but could improve in mentoring junior team members." },
    { "FEEDBACK_DATE": "2022-03-12", "FEEDBACK_TEXT": "Now actively mentors junior team members and is a valuable resource for their development." }
  ] },

  { "EMPLOYEE_ID": 42, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-03-13", "FEEDBACK_TEXT": "Employee needs to work on public speaking and presenting in front of large audiences." },
    { "FEEDBACK_DATE": "2022-03-13", "FEEDBACK_TEXT": "Has become an effective public speaker and now leads presentations with confidence." }
  ] },

  { "EMPLOYEE_ID": 43, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-03-14", "FEEDBACK_TEXT": "This employee needs to improve on providing timely updates and communication with the team." },
    { "FEEDBACK_DATE": "2022-03-14", "FEEDBACK_TEXT": "Has improved significantly in communication and consistently provides timely project updates." }
  ] },

  { "EMPLOYEE_ID": 44, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-03-15", "FEEDBACK_TEXT": "Employee is a valuable team member but struggles with prioritizing multiple competing deadlines." },
    { "FEEDBACK_DATE": "2022-03-15", "FEEDBACK_TEXT": "Now effectively prioritizes tasks and meets competing deadlines with ease." }
  ] },

  { "EMPLOYEE_ID": 45, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-03-16", "FEEDBACK_TEXT": "Employee is reliable but could benefit from more initiative in suggesting process improvements." },
    { "FEEDBACK_DATE": "2023-03-16", "FEEDBACK_TEXT": "Has taken the initiative in suggesting process improvements and is more proactive in problem-solving." }
  ] },

  { "EMPLOYEE_ID": 46, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-03-17", "FEEDBACK_TEXT": "Employee is good at handling day-to-day tasks but struggles with leading larger projects." },
    { "FEEDBACK_DATE": "2023-03-17", "FEEDBACK_TEXT": "Has taken on larger projects and now leads them successfully, driving key initiatives forward." }
  ] },

  { "EMPLOYEE_ID": 47, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-03-18", "FEEDBACK_TEXT": "Employee is skilled but could benefit from greater attention to detail in their work." },
    { "FEEDBACK_DATE": "2022-03-18", "FEEDBACK_TEXT": "Has improved attention to detail and is now known for delivering error-free work." }
  ] },

  { "EMPLOYEE_ID": 48, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-03-19", "FEEDBACK_TEXT": "Employee is a strong technical contributor but could work on balancing their workload better." },
    { "FEEDBACK_DATE": "2023-03-19", "FEEDBACK_TEXT": "Now effectively balances workload and contributes to multiple high-priority projects." }
  ] },

  { "EMPLOYEE_ID": 49, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-03-20", "FEEDBACK_TEXT": "This employee needs to improve their approach to conflict resolution within teams." },
    { "FEEDBACK_DATE": "2022-03-20", "FEEDBACK_TEXT": "Has made significant improvements in conflict resolution and maintains harmony within the team." }
  ] },

  { "EMPLOYEE_ID": 50, "FEEDBACK": [
    { "FEEDBACK_DATE": "2020-03-21", "FEEDBACK_TEXT": "Employee is dependable but needs to take more initiative in seeking out opportunities for growth." },
    { "FEEDBACK_DATE": "2023-03-21", "FEEDBACK_TEXT": "Has become more proactive in seeking growth opportunities and now contributes to organizational development." }
  ] }


]

feedback_col = employee_feedback_col.insert_many(feedback_list)


#print list of the _id values of the inserted documents:
print(feedback_col.inserted_ids)

Offboarding_review_list = [
  { 
    "EMPLOYEE_ID": 12, 
    "OFFBOARDING_DATE": "2023-10-14", 
    "REASON_FOR_LEAVING": "Termination for poor performance", 
    "OFFBOARDING_DECISION_CATEGORY": "Terminated", 
    "FEEDBACK_TEXT": "Employee was terminated due to consistent underperformance, failing to meet key performance indicators despite multiple performance reviews and improvement plans.",
    "EXIT_INTERVIEW_FEEDBACK": "Employee expressed frustration with unclear expectations and lack of support in role. They indicated feeling overwhelmed by workload and suggested that better communication and resources could have helped them succeed."
  },
  { 
    "EMPLOYEE_ID": 22, 
    "OFFBOARDING_DATE": "2023-11-03", 
    "REASON_FOR_LEAVING": "Termination due to behavior issues", 
    "OFFBOARDING_DECISION_CATEGORY": "Terminated", 
    "FEEDBACK_TEXT": "Employee was terminated due to behavior issues, including inappropriate conduct towards colleagues and a failure to adhere to company values after multiple warnings.",
    "EXIT_INTERVIEW_FEEDBACK": "Employee claimed that they felt misjudged by management and that their behavior was misunderstood. They mentioned issues with team dynamics and felt that their grievances were not taken seriously, but ultimately accepted the termination decision."
  },
  { 
    "EMPLOYEE_ID": 15, 
    "OFFBOARDING_DECISION_CATEGORY": "Under Review", 
    "FEEDBACK_TEXT": "Employee has had multiple performance issues but is currently being monitored for improvement. Further review is required to determine whether termination or a different course of action will be taken."
  },
  { 
    "EMPLOYEE_ID": 18,  
    "OFFBOARDING_DECISION_CATEGORY": "Under Review", 
    "FEEDBACK_TEXT": "Employee has received consistent feedback about performance, but no final decision has been made. Further evaluations are needed before making an offboarding decision."
  },
  { 
    "EMPLOYEE_ID": 10, 
    "OFFBOARDING_DECISION_CATEGORY": "Under Review", 
    "FEEDBACK_TEXT": "Employee has a strong performance record, but there have been some recent issues with team collaboration. It's under review whether additional support and mentoring could resolve these issues."
  },
  { 
    "EMPLOYEE_ID": 5, 
    "OFFBOARDING_DECISION_CATEGORY": "Under Review", 
    "FEEDBACK_TEXT": "Employee has shown inconsistencies in meeting expectations. While there have been some positive contributions, the team dynamics are being monitored, and the employee's future status is still under consideration."
  },
  { 
    "EMPLOYEE_ID": 20, 
    "OFFBOARDING_DECISION_CATEGORY": "Under Review", 
    "FEEDBACK_TEXT": "Employee has had a few performance-related concerns but is currently undergoing a performance improvement plan. A decision will be made based on progress after the review period."
  },
  { 
    "EMPLOYEE_ID": 25, 
    "OFFBOARDING_DECISION_CATEGORY": "Under Review", 
    "FEEDBACK_TEXT": "Employee's recent behavior and performance have raised concerns. There is currently a review process in place to determine if corrective action or potential termination is the appropriate course of action."
  }
]

offboarding_col = offboarding_review_col.insert_many(Offboarding_review_list)

employee_development_list = [
  { 
    "EMPLOYEE_ID": 15, 
    "DEVELOPMENT_PROGRAMS": [
      {
        "PROGRAM_NAME": "Advanced Data Analytics",
        "START_DATE": "2023-01-15",
        "END_DATE": "2023-04-15",
        "STATUS": "Completed",
        "SKILLS_ACQUIRED": ["Data Visualization", "Predictive Modeling", "SQL"],
        "IMPACT": "Improved data analysis capabilities, leading to more efficient decision-making."
      },
      {
        "PROGRAM_NAME": "Leadership Development Program",
        "START_DATE": "2024-06-01",
        "END_DATE": "2024-09-01",
        "STATUS": "In Progress",
        "SKILLS_ACQUIRED": ["Team Management", "Conflict Resolution", "Strategic Planning"],
        "IMPACT": "Developing leadership skills for future management roles."
      }
    ],
    "PERFORMANCE_GOALS": [
      {
        "GOAL": "Increase Sales Conversion Rate",
        "TARGET_DATE": "2024-12-31",
        "PROGRESS": "50%",
        "COMMENTS": "Working on improving negotiation skills with training."
      },
      {
        "GOAL": "Improve Customer Satisfaction Scores",
        "TARGET_DATE": "2024-09-30",
        "PROGRESS": "70%",
        "COMMENTS": "Positive results from customer feedback sessions, but more work needed on follow-ups."
      }
    ],
    "AREAS_FOR_IMPROVEMENT": [
      "Time Management",
      "Conflict Resolution",
      "Presentation Skills"
    ]
  },
  { 
    "EMPLOYEE_ID": 18, 
    "DEVELOPMENT_PROGRAMS": [
      {
        "PROGRAM_NAME": "Project Management Fundamentals",
        "START_DATE": "2023-03-01",
        "END_DATE": "2023-06-01",
        "STATUS": "Completed",
        "SKILLS_ACQUIRED": ["Project Planning", "Risk Management", "Budget Management"],
        "IMPACT": "Successfully managed smaller projects, improving efficiency."
      },
      {
        "PROGRAM_NAME": "Effective Communication Workshop",
        "START_DATE": "2024-02-01",
        "END_DATE": "2024-02-28",
        "STATUS": "Completed",
        "SKILLS_ACQUIRED": ["Public Speaking", "Written Communication"],
        "IMPACT": "Improved communication with cross-functional teams."
      }
    ],
    "PERFORMANCE_GOALS": [
      {
        "GOAL": "Lead Cross-Departmental Projects",
        "TARGET_DATE": "2024-12-31",
        "PROGRESS": "30%",
        "COMMENTS": "Attending project management training to prepare."
      },
      {
        "GOAL": "Increase Team Collaboration",
        "TARGET_DATE": "2024-09-30",
        "PROGRESS": "50%",
        "COMMENTS": "Initiated weekly team meetings, but results are still being evaluated."
      }
    ],
    "AREAS_FOR_IMPROVEMENT": [
      "Delegation Skills",
      "Time Management",
      "Confidence in Decision-Making"
    ]
  },
  { 
    "EMPLOYEE_ID": 10, 
    "DEVELOPMENT_PROGRAMS": [
      {
        "PROGRAM_NAME": "Advanced Marketing Strategies",
        "START_DATE": "2023-05-01",
        "END_DATE": "2023-08-01",
        "STATUS": "Completed",
        "SKILLS_ACQUIRED": ["Market Analysis", "Digital Advertising", "SEO"],
        "IMPACT": "Contributed to increased brand visibility and sales."
      },
      {
        "PROGRAM_NAME": "Creative Problem-Solving Techniques",
        "START_DATE": "2024-03-01",
        "END_DATE": "2024-05-31",
        "STATUS": "In Progress",
        "SKILLS_ACQUIRED": ["Creative Thinking", "Innovation", "Critical Thinking"],
        "IMPACT": "Developing new approaches to solve complex marketing challenges."
      }
    ],
    "PERFORMANCE_GOALS": [
      {
        "GOAL": "Enhance Online Marketing Strategies",
        "TARGET_DATE": "2024-12-31",
        "PROGRESS": "60%",
        "COMMENTS": "Working on enhancing PPC campaigns to improve conversion rates."
      },
      {
        "GOAL": "Expand Social Media Reach",
        "TARGET_DATE": "2024-09-30",
        "PROGRESS": "40%",
        "COMMENTS": "Collaborating with the creative team on new content formats."
      }
    ],
    "AREAS_FOR_IMPROVEMENT": [
      "Data Interpretation",
      "Client Relationship Management",
      "Team Leadership"
    ]
  },
  { 
    "EMPLOYEE_ID": 5, 
    "DEVELOPMENT_PROGRAMS": [
      {
        "PROGRAM_NAME": "Financial Analysis Mastery",
        "START_DATE": "2023-07-01",
        "END_DATE": "2023-10-01",
        "STATUS": "Completed",
        "SKILLS_ACQUIRED": ["Financial Forecasting", "Budget Planning", "Accounting Software"],
        "IMPACT": "Increased accuracy in financial reports and forecasting."
      },
      {
        "PROGRAM_NAME": "Negotiation Skills Training",
        "START_DATE": "2024-04-01",
        "END_DATE": "2024-04-30",
        "STATUS": "Upcoming",
        "SKILLS_ACQUIRED": ["Negotiation", "Conflict Resolution"],
        "IMPACT": "Preparing for more client-facing roles."
      }
    ],
    "PERFORMANCE_GOALS": [
      {
        "GOAL": "Reduce Operational Costs by 15%",
        "TARGET_DATE": "2024-12-31",
        "PROGRESS": "35%",
        "COMMENTS": "Tracking current expenditures and evaluating cost-saving strategies."
      },
      {
        "GOAL": "Improve Financial Reporting Timeliness",
        "TARGET_DATE": "2024-06-30",
        "PROGRESS": "80%",
        "COMMENTS": "Made significant progress in streamlining reporting processes."
      }
    ],
    "AREAS_FOR_IMPROVEMENT": [
      "Team Collaboration",
      "Confidence in Leadership",
      "Presentation Skills"
    ]
  },
  { 
    "EMPLOYEE_ID": 20, 
    "DEVELOPMENT_PROGRAMS": [
      {
        "PROGRAM_NAME": "Software Development Bootcamp",
        "START_DATE": "2023-06-01",
        "END_DATE": "2023-09-01",
        "STATUS": "Completed",
        "SKILLS_ACQUIRED": ["Python Programming", "Database Management", "Software Development Lifecycle"],
        "IMPACT": "Improved coding efficiency and ability to handle complex development projects."
      },
      {
        "PROGRAM_NAME": "Agile Project Management",
        "START_DATE": "2024-05-01",
        "END_DATE": "2024-07-01",
        "STATUS": "Upcoming",
        "SKILLS_ACQUIRED": ["Agile Methodology", "Scrum Master", "Team Leadership"],
        "IMPACT": "Preparing for a transition into a more senior role with project management responsibilities."
      }
    ],
    "PERFORMANCE_GOALS": [
      {
        "GOAL": "Complete Development of New Internal Tools",
        "TARGET_DATE": "2024-12-31",
        "PROGRESS": "25%",
        "COMMENTS": "Working closely with the product team to finalize requirements and design."
      },
      {
        "GOAL": "Improve Code Quality Standards",
        "TARGET_DATE": "2024-09-30",
        "PROGRESS": "50%",
        "COMMENTS": "Participating in code reviews to ensure high-quality deliverables."
      }
    ],
    "AREAS_FOR_IMPROVEMENT": [
      "Code Optimization",
      "Mentoring Junior Developers",
      "Documentation"
    ]
  },
  { 
    "EMPLOYEE_ID": 25, 
    "DEVELOPMENT_PROGRAMS": [
      {
        "PROGRAM_NAME": "Business Analysis Fundamentals",
        "START_DATE": "2023-02-01",
        "END_DATE": "2023-05-01",
        "STATUS": "Completed",
        "SKILLS_ACQUIRED": ["Business Process Mapping", "Requirements Gathering", "Risk Analysis"],
        "IMPACT": "Strengthened ability to analyze and improve business processes."
      },
      {
        "PROGRAM_NAME": "Negotiation and Influencing Skills",
        "START_DATE": "2024-07-01",
        "END_DATE": "2024-09-01",
        "STATUS": "Upcoming",
        "SKILLS_ACQUIRED": ["Negotiation", "Influencing Others", "Stakeholder Management"],
        "IMPACT": "Preparing for more strategic roles in management."
      }
    ],
    "PERFORMANCE_GOALS": [
      {
        "GOAL": "Enhance Stakeholder Engagement",
        "TARGET_DATE": "2024-12-31",
        "PROGRESS": "40%",
        "COMMENTS": "Building stronger relationships with key stakeholders across departments."
      },
      {
        "GOAL": "Improve Risk Management Practices",
        "TARGET_DATE": "2024-09-30",
        "PROGRESS": "60%",
        "COMMENTS": "Working closely with risk management team to implement better strategies."
      }
    ],
    "AREAS_FOR_IMPROVEMENT": [
      "Decision-Making",
      "Stakeholder Communication",
      "Time Management"
    ]
  }
 
]

employee_personal_development_col = employee_development_col.insert_many(employee_development_list)

remprimands_list = [
  {
    "EMPLOYEE_ID": 12,
    "REPRIMAND_DATE": "2023-09-10",
    "REPRIMAND_TYPE": "Performance Issue",
    "REASON": "Consistently failing to meet KPIs despite repeated performance reviews and corrective action plans.",
    "ACTION_TAKEN": "Written warning issued. Employee to attend performance improvement plan (PIP) sessions.",
    "FOLLOW_UP_DATE": "2023-09-24",
    "HR_INVOLVEMENT": True
  },
  {
    "EMPLOYEE_ID": 22,
    "REPRIMAND_DATE": "2023-11-03",
    "REPRIMAND_TYPE": "Behavioral Issue",
    "REASON": "Inappropriate conduct towards colleagues, including hostile language during meetings and refusal to follow team norms.",
    "ACTION_TAKEN": "Final written warning issued. Employee advised to attend conflict resolution training.",
    "FOLLOW_UP_DATE": "2023-11-17",
    "HR_INVOLVEMENT": True
  },
  {
    "EMPLOYEE_ID": 15,
    "REPRIMAND_DATE": "2023-07-12",
    "REPRIMAND_TYPE": "Attendance Issue",
    "REASON": "Frequent unexcused absences, affecting project deadlines and team productivity.",
    "ACTION_TAKEN": "Verbal warning issued. Employee required to submit a monthly attendance report.",
    "FOLLOW_UP_DATE": "2023-07-26",
    "HR_INVOLVEMENT": False
  },
  {
    "EMPLOYEE_ID": 18,
    "REPRIMAND_DATE": "2023-08-14",
    "REPRIMAND_TYPE": "Performance Issue",
    "REASON": "Failure to meet quarterly performance targets despite several coaching sessions.",
    "ACTION_TAKEN": "Written warning issued. Employee to complete performance review and attend additional training.",
    "FOLLOW_UP_DATE": "2023-09-05",
    "HR_INVOLVEMENT": True
  },
  {
    "EMPLOYEE_ID": 10,
    "REPRIMAND_DATE": "2023-06-20",
    "REPRIMAND_TYPE": "Behavioral Issue",
    "REASON": "Disrespectful tone in team meetings, creating a hostile environment.",
    "ACTION_TAKEN": "Verbal warning issued. Employee required to attend professional conduct workshop.",
    "FOLLOW_UP_DATE": "2023-07-10",
    "HR_INVOLVEMENT": False
  },
  {
    "EMPLOYEE_ID": 5,
    "REPRIMAND_DATE": "2023-07-18",
    "REPRIMAND_TYPE": "Attendance Issue",
    "REASON": "Chronic tardiness, impacting team productivity.",
    "ACTION_TAKEN": "Written warning issued. Employee to meet with HR to discuss attendance improvement plan.",
    "FOLLOW_UP_DATE": "2023-08-01",
    "HR_INVOLVEMENT": True
  },
  {
    "EMPLOYEE_ID": 20,
    "REPRIMAND_DATE": "2023-09-05",
    "REPRIMAND_TYPE": "Performance Issue",
    "REASON": "Failure to deliver key project milestones on time, causing delays in the team’s schedule.",
    "ACTION_TAKEN": "Final written warning issued. Employee to attend additional project management training and submit progress reports.",
    "FOLLOW_UP_DATE": "2023-09-19",
    "HR_INVOLVEMENT": True
  },
  {
    "EMPLOYEE_ID": 25,
    "REPRIMAND_DATE": "2023-06-30",
    "REPRIMAND_TYPE": "Behavioral Issue",
    "REASON": "Inappropriate communication with team members during project discussions, leading to team conflict.",
    "ACTION_TAKEN": "Verbal warning issued. Employee to complete team-building and communication skills training.",
    "FOLLOW_UP_DATE": "2023-07-14",
    "HR_INVOLVEMENT": False
  },
  {
    "EMPLOYEE_ID": 1,
    "REPRIMAND_DATE": "2023-08-05",
    "REPRIMAND_TYPE": "Performance Issue",
    "REASON": "Employee failed to meet monthly targets and did not respond promptly to feedback from the manager.",
    "ACTION_TAKEN": "Written warning issued. Employee required to attend a performance improvement plan (PIP) session.",
    "FOLLOW_UP_DATE": "2023-08-19",
    "HR_INVOLVEMENT": True
  },  
  {
    "EMPLOYEE_ID": 3,
    "REPRIMAND_DATE": "2023-09-10",
    "REPRIMAND_TYPE": "Attendance Issue",
    "REASON": "Employee was repeatedly late to work without valid reasons, causing disruption to team workflow.",
    "ACTION_TAKEN": "Written warning issued. Employee instructed to follow the company's attendance policy.",
    "FOLLOW_UP_DATE": "2023-09-24",
    "HR_INVOLVEMENT": False
  },
  {
    "EMPLOYEE_ID": 8,
    "REPRIMAND_DATE": "2023-09-01",
    "REPRIMAND_TYPE": "Attendance Issue",
    "REASON": "Employee has missed several workdays without providing valid justification or advance notice.",
    "ACTION_TAKEN": "Final written warning issued. Employee reminded of the attendance policy and expected to maintain punctuality.",
    "FOLLOW_UP_DATE": "2023-09-15",
    "HR_INVOLVEMENT": True
  }

]

reprimand_col = reprimands_col.insert_many(remprimands_list)
print(mydb.list_collection_names())
#Show documents in collection "customersTest"
myresult = employee_feedback_col.find()

#print the result:
for x in myresult:
    print(x)


# Queries

# Query for terminated employees and convert to a list
terminated_employees = list(offboarding_review_col.find({
    "OFFBOARDING_DECISION_CATEGORY": "Terminated"
}))

print(terminated_employees)

# Find feedback for terminated employees 
for x in employee_feedback_col.find({"EMPLOYEE_ID": {"$in": [12, 22]}}):
    print(x)
under_review_employees = list(offboarding_review_col.find({
    "OFFBOARDING_DECISION_CATEGORY": "Under Review"
}))
print(under_review_employees)


#Sorting employees by most recent reprimands
mydoc = reprimands_col.find().sort("REPRIMAND_DATE", -1)
for x in mydoc:
     print(x)


# Identifying employees with more serious remprimands where HR had to be involved 
for reprimand_HR in reprimands_col.find({"HR_INVOLVEMENT": True}):
    print(reprimand_HR)

# Find key areas for improvement within the firm
for x in employee_development_col.find({},{ "EMPLOYEE_ID":1, "AREAS_FOR_IMPROVEMENT":1}):
    print(x)


# Updating
# Update the employee 20 offboarding decision to terminated
myquery = { "EMPLOYEE_ID": 20 }
newvalues = { "$set": { "OFFBOARDING_DECISION_CATEGORY": "Terminated" } }

result = offboarding_review_col.update_one(myquery, newvalues)

# Print the result of the update
print(f"Documents matched: {result.matched_count}")
print(f"Documents modified: {result.modified_count}")

# Print documents in the collection after the update
for x in offboarding_review_col.find({"EMPLOYEE_ID": 20}):
  print(x)

# Deleting
# Removing employee 10 from offboarding review as deemd valuable
offboarding_review_col.delete_one({"EMPLOYEE_ID": 10})


#Visualisations
# import matplotlib.pyplot as plt
# from pymongo import MongoClient
# from collections import Counter
# # Fetch the reprimands from the collection and count by REPRIMAND_TYPE
# reprimand_types = [reprimand['REPRIMAND_TYPE'] for reprimand in reprimands_col.find()]

# # Use Counter to get the frequency of each reprimand type
# reprimand_counts = Counter(reprimand_types)

# # Extract the reprimand types and counts
# labels = list(reprimand_counts.keys())
# values = list(reprimand_counts.values())

# # Plot the bar graph
# plt.figure(figsize=(10, 6))  # Optional: Adjust the size of the plot
# plt.bar(labels, values, color='skyblue')
# plt.xlabel('Reprimand Type')
# plt.ylabel('Count')
# plt.title('Count of Reprimands by Type')
# plt.xticks(rotation=45)  # Rotate the x labels for better visibility
# plt.show()
# from textblob import TextBlob
# import matplotlib.pyplot as plt

# # Function to get sentiment polarity (positive/negative score)
# def get_sentiment(text):
#     blob = TextBlob(text)
#     return blob.sentiment.polarity

# # Let's assume `feedback_list` is the list containing employee feedback.
# employee_sentiments = {}

# for employee in feedback_list:
#     sentiments = [get_sentiment(feedback['FEEDBACK_TEXT']) for feedback in employee['FEEDBACK']]
#     avg_sentiment = sum(sentiments) / len(sentiments)  # average sentiment score

#     # Classify employee as positive, negative, or neutral based on average sentiment
#     if avg_sentiment > 0.1:
#         employee_sentiments[employee['EMPLOYEE_ID']] = 'Positive'
#     elif avg_sentiment < -0.1:
#         employee_sentiments[employee['EMPLOYEE_ID']] = 'Negative'
#     else:
#         employee_sentiments[employee['EMPLOYEE_ID']] = 'Neutral'

# # Visualize the distribution of sentiment categories
# sentiment_counts = {'Positive': 0, 'Negative': 0, 'Neutral': 0}
# for sentiment in employee_sentiments.values():
#     sentiment_counts[sentiment] += 1

# # Create a bar chart to show the distribution
# plt.bar(sentiment_counts.keys(), sentiment_counts.values())
# plt.xlabel('Sentiment')
# plt.ylabel('Number of Employees')
# plt.title('Employee Feedback Sentiment Distribution')
# plt.show()
