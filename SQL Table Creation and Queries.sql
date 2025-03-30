
-- Step 1: Creating the EMPLOYEE table without foreign keys  
DROP TABLE IF EXISTS [EmployeeDatas].EMPLOYEE;  

 
CREATE TABLE [EmployeeDatas].EMPLOYEE(  
    employeeID INT PRIMARY KEY,  
    firstName VARCHAR(50) NOT NULL,  
    lastName VARCHAR(50) NOT NULL,  
    email VARCHAR(100) UNIQUE NOT NULL,  
    hireDate DATE NOT NULL,  
    departmentID INT,  -- FK added later  
    billRate FLOAT CHECK (billRate >= 0),  
    mentorID INT,      -- FK added later  
    atRiskStatus BIT NOT NULL,  
    payrollID INT,      -- FK added later  
    performanceID INT  -- FK added later  
);  

--step 2: insert the employee data  

INSERT INTO [EmployeeDatas].[EMPLOYEE] (employeeID, firstName, lastName, email, hireDate, departmentID, billRate, mentorID, atRiskStatus, payrollID, performanceID)   
VALUES   
(1, 'John', 'Doe', 'johndoe@consultancy.com', '2015-01-10', 1, 200, NULL, 0, 1, 1),   
(2, 'Jane', 'Smith', 'janesmith@consultancy.com', '2020-06-10', 2, 145, 1, 1, 2, 2),   
(3, 'Mark', 'Johnson', 'markjohnson@consultancy.com', '2017-03-25', 2, 170, 1, 0, 3, 3),   
(4, 'Emily', 'Davis', 'emilydavis@consultancy.com', '2021-07-01', 3, 115, NULL, 1, 4, 4),   
(5, 'Michael', 'Brown', 'michaelbrown@consultancy.com', '2018-02-20', 1, 180, 1, 0, 5, 5),   
(6, 'Samantha', 'Clark', 'samanthaclark@consultancy.com', '2020-08-10', 4, 120, NULL, 1, 6, 6),   
(7, 'David', 'Lewis', 'davidlewis@consultancy.com', '2015-04-20', 5, 180, NULL, 0, 7, 7),   
(8, 'Rachel', 'White', 'rachelwhite@consultancy.com', '2021-02-25', 3, 110, 7, 1, 8, 8),   
(9, 'George', 'Miller', 'georgemiller@consultancy.com', '2019-06-30', 2, 150, 1, 0, 9, 9),   
(10, 'Linda', 'Wilson', 'lindawilson@consultancy.com', '2016-11-15', 1, 170, 1, 1, 10, 10),   
(11, 'Henry', 'Anderson', 'henryanderson@consultancy.com', '2020-05-10', 4, 135, NULL, 0, 11, 11),   
(12, 'Paul', 'Martinez', 'paulmartinez@consultancy.com', '2018-03-15', 5, 150, 7, 0, 12, 12),   
(13, 'Susan', 'Taylor', 'susantaylor@consultancy.com', '2021-04-11', 2, 120, 3, 1, 13, 13),   
(14, 'Daniel', 'Thomas', 'danielthomas@consultancy.com', '2017-06-10', 1, 175, 1, 0, 14, 14),   
(15, 'Patricia', 'Jackson', 'patriciajackson@consultancy.com', '2021-01-01', 3, 110, NULL, 1, 15, 15),   
(16, 'James', 'Harris', 'jamesharris@consultancy.com', '2016-09-10', 4, 160, 7, 0, 16, 16),   
(17, 'Deborah', 'Clark', 'deborahclark@consultancy.com', '2020-11-25', 5, 125, NULL, 1, 17, 17),   
(18, 'Samuel', 'Lewis', 'samuellewis@consultancy.com', '2017-12-15', 1, 145, 5, 0, 18, 18),   
(19, 'Natalie', 'Roberts', 'natalieroberts@consultancy.com', '2019-09-20', 2, 130, 3, 1, 19, 19),   
(20, 'Alan', 'Scott', 'alanscott@consultancy.com', '2018-06-30', 3, 155, 1, 0, 20, 20),   
(21, 'Charlotte', 'Walker', 'charlottewalker@consultancy.com', '2020-03-15', 5, 140, 7, 1, 21, 21),   
(22, 'Matthew', 'Allen', 'matthewallen@consultancy.com', '2021-01-25', 4, 125, NULL, 0, 22, 22),   
(23, 'Sophie', 'King', 'sophieking@consultancy.com', '2021-02-10', 2, 115, NULL, 1, 23, 23),   
(24, 'Peter', 'Lopez', 'peterlopez@consultancy.com', '2018-07-01', 1, 160, 1, 0, 24, 24),   
(25, 'Jessica', 'Hill', 'jessicahill@consultancy.com', '2017-01-17', 3, 135, 4, 1, 25, 25),   
(26, 'Oscar', 'Wright', 'oscarwright@consultancy.com', '2019-08-12', 5, 140, 7, 0, 26, 26),   
(27, 'Olivia', 'Green', 'oliviagreen@consultancy.com', '2020-11-25', 1, 120, NULL, 1, 27, 27),   
(28, 'Jacob', 'Young', 'jacobyoung@consultancy.com', '2020-07-10', 4, 125, 7, 0, 28, 28),   
(29, 'Aiden', 'Baker', 'aidenbaker@consultancy.com', '2016-06-17', 2, 150, 10, 0, 29, 29),   
(30, 'Grace', 'Adams', 'graceadams@consultancy.com', '2021-02-01', 3, 110, 7, 1, 30, 30),   
(31, 'Jack', 'Nelson', 'jacknelson@consultancy.com', '2020-09-21', 4, 130, NULL, 0, 31, 31),   
(32, 'Victoria', 'Carter', 'victoriacarter@consultancy.com', '2020-04-15', 1, 140, 1, 1, 32, 32),   
(33, 'Anthony', 'Perez', 'anthonyperez@consultancy.com', '2016-03-11', 5, 150, 7, 0, 33, 33),   
(34, 'Isabella', 'Morris', 'isabellamorris@consultancy.com', '2021-03-05', 2, 115, NULL, 1, 34, 34),   
(35, 'Daniel', 'Miller', 'danielmiller@consultancy.com', '2021-02-12', 3, 130, 4, 1, 35, 35),   
(36, 'Benjamin', 'Roberts', 'benjaminroberts@consultancy.com', '2017-08-23', 4, 150, 7, 0, 36, 36),   
(37, 'Ella', 'Turner', 'ellaturner@consultancy.com', '2020-10-01', 1, 145, NULL, 0, 37, 37),   
(38, 'Lucas', 'Green', 'lucasgreen@consultancy.com', '2018-01-20', 5, 160, 12, 0, 38, 38),   
(39, 'Megan', 'Mitchell', 'meganmitchell@consultancy.com', '2021-01-07', 3, 120, 1, 1, 39, 39),   
(40, 'Lucas', 'Parker', 'lucasparker@consultancy.com', '2016-12-03', 2, 135, NULL, 1, 40, 40),   
(41, 'Leo', 'Hughes', 'leohughes@consultancy.com', '2017-10-01', 1, 155, 1, 0, 41, 41),   
(42, 'Chloe', 'Perez', 'chloeperez@consultancy.com', '2020-08-25', 4, 120, 7, 1, 42, 42),   
(43, 'Zachary', 'Young', 'zacharyyoung@consultancy.com', '2021-03-15', 2, 130, NULL, 0, 43, 43),   
(44, 'Ethan', 'King', 'ethanking@consultancy.com', '2019-06-07', 5, 150, 7, 0, 44, 44),   
(45, 'Mia', 'Lopez', 'mialopez@consultancy.com', '2020-04-05', 3, 115, 1, 1, 45, 45),   
(46, 'Ryan', 'Lopez', 'ryanlopez@consultancy.com', '2016-09-12', 1, 160, 13, 0, 46, 46),   
(47, 'Evelyn', 'Walker', 'evelynwalker@consultancy.com', '2018-05-14', 4, 130, 1, 1, 47, 47),   
(48, 'Oliver', 'Martin', 'olivermartin@consultancy.com', '2019-04-15', 2, 145, 5, 0, 48, 48),   
(49, 'Sophia', 'Rodriguez', 'sophiarodriguez@consultancy.com', '2021-04-08', 5, 120, NULL, 1, 49, 49),   
(50, 'Ryan', 'Foster', 'ryanfoster@consultancy.com', '2020-12-25', 3, 145, 6, 0, 50, 50);   
 

 -- Step 3: Creating all other tables with their foreign keys defined  

 CREATE TABLE [EmployeeDatas].[DEPARTMENT] (  
    departmentID INT PRIMARY KEY,  
    departmentName VARCHAR(100) NOT NULL,  
    headID INT UNIQUE, --   
    annualTarget FLOAT CHECK (annualTarget >= 0),  
    departmentBudget FLOAT CHECK (departmentBudget >= 0),  
    currentRevenue FLOAT CHECK (currentRevenue >= 0),  
    headcount INT CHECK (headcount >= 0),  
    specialization VARCHAR(100),  
    FOREIGN KEY (headID) REFERENCES [EmployeeDatas].[EMPLOYEE](employeeID)  

);  
 

CREATE TABLE [EmployeeDatas].[PAYROLL] (  

    payrollID INT PRIMARY KEY,  
    employeeID INT,  
    salary FLOAT CHECK (salary >= 0),  
    payPeriodStart DATE NOT NULL,  
    payPeriodEnd DATE NOT NULL,  
    overtimePay FLOAT CHECK (overtimePay >= 0),  
    bonus FLOAT CHECK (bonus >= 0),  
    commission FLOAT CHECK (commission >= 0),  
    billableHours FLOAT CHECK (billableHours >= 0),  
    nonBillableHours FLOAT CHECK (nonBillableHours >= 0),  
    benefits FLOAT CHECK (benefits >= 0),  
    totalDeductions FLOAT CHECK (totalDeductions >= 0),  
    netPay FLOAT CHECK (netPay >= 0),  
    paymentDate DATE NOT NULL,  
    paymentStatus VARCHAR(50),  
    finalPaycheck BIT,  
    severanceIncluded BIT,  
    FOREIGN KEY (employeeID) REFERENCES [EmployeeDatas].[EMPLOYEE](employeeID)  

);  


CREATE TABLE [EmployeeDatas].[REASON_CATEGORY] (  

    reasonCategoryID INT PRIMARY KEY,  
    reasonName VARCHAR(100) NOT NULL  

);  

CREATE TABLE [EmployeeDatas].[OFFBOARDING_DECISION] (  

    decisionID INT PRIMARY KEY,  
    employeeID INT,  
    reviewDate DATE,  
    decisionMaker INT,  
    decisionStatus VARCHAR(50),  
    terminationDate DATE,  
    severanceAmount FLOAT CHECK (severanceAmount >= 0),  
    reasonCategoryID INT,  
    performanceScore FLOAT CHECK (performanceScore >= 0 AND performanceScore <= 100),  
    replacementNeeded BIT,  
    exitInterviewScheduled BIT,  
    legalReviewComplete BIT,  
    appealStatus VARCHAR(50),  
    FOREIGN KEY (employeeID) REFERENCES [EmployeeDatas].[EMPLOYEE](employeeID),  
    FOREIGN KEY (decisionMaker) REFERENCES [EmployeeDatas].[DEPARTMENT](headID),  
    FOREIGN KEY (reasonCategoryID) REFERENCES [EmployeeDatas].[REASON_CATEGORY](reasonCategoryID)  

);  
 

CREATE TABLE [EmployeeDatas].[PERFORMANCE_DATA] (  

    performanceID INT PRIMARY KEY,  
    employeeID INT,  
    evaluationDate DATE NOT NULL,  
    utilizationRate FLOAT CHECK (utilizationRate >= 0 AND utilizationRate <= 100),  
    clientFeedbackScore INT CHECK (clientFeedbackScore >= 0 AND clientFeedbackScore <= 5),  
    revenueGenerated FLOAT CHECK (revenueGenerated >= 0),  
    teamworkRating INT CHECK (teamworkRating >= 0 AND teamworkRating <= 5),  
    performanceRating INT CHECK (performanceRating >= 0 AND performanceRating <= 5),  
    FOREIGN KEY (employeeID) REFERENCES [EmployeeDatas].[EMPLOYEE](employeeID)  

);  

CREATE TABLE [EmployeeDatas].[SKILL_COMPENDIUM] (  

    skill_id INT IDENTITY(1,1) PRIMARY KEY,  
    skill_name VARCHAR(50) NOT NULL,  
    skill_category VARCHAR(50),  
    skill_description VARCHAR(250),  
    created_at DATETIME2 DEFAULT GETDATE(),  
    updated_at DATETIME2 DEFAULT GETDATE()  

);  
 

CREATE TABLE [EmployeeDatas].[EMPLOYEE_SKILLS] (  

    employee_skills_id INT NOT NULL,  
    employee_id INT NOT NULL,  
    skill_id INT NOT NULL,  
    proficiency_level VARCHAR(40),  
    last_used DATE,  
    PRIMARY KEY (employee_id, skill_id),  
    FOREIGN KEY (employee_id) REFERENCES [EmployeeDatas].[EMPLOYEE](employeeID),  
    FOREIGN KEY (skill_id) REFERENCES [EmployeeDatas].[SKILL_COMPENDIUM](skill_id)  

);  


-- Step 4: Insert data into Department table   

INSERT INTO [EmployeeDatas].[DEPARTMENT] (departmentID, departmentName, headID, annualTarget, departmentBudget, currentRevenue, headcount, specialization)   

VALUES   
(1, 'Strategy', 1, 5000000, 4800000, 4000000, 10, 'Strategic Consulting'),   
(2, 'Technology', 5, 7000000, 6800000, 5500000, 12, 'IT Consulting'),   
(3, 'Operations', 14, 4500000, 4300000, 3600000, 8, 'Operations Optimization'),   
(4, 'HR', 16, 3000000, 2900000, 0, 3200000, 'Human Resources Consulting'),   
(5, 'Finance', 7, 6000000, 5900000, 6100000, 7, 'Financial Consulting');   



  -- step 5: Insert the payroll data   

INSERT INTO [EmployeeDatas].[PAYROLL] (payrollID, employeeID, salary, payPeriodStart, payPeriodEnd, overtimePay, bonus, commission, billableHours, nonBillableHours, benefits, totalDeductions, netPay, paymentDate, paymentStatus, finalPaycheck, severanceIncluded)   

VALUES   
(1, 1, 120000.00, '2024-01-01', '2024-01-31', 500, 2000, 1000, 160, 20, 500, 300, 9500, '2024-02-05', 'Completed', 0, 0),   
(2, 2, 90000.00, '2024-01-01', '2024-01-31', 300, 1500, 800, 150, 30, 400, 250, 7700, '2024-02-05', 'Completed', 0, 0),   
(3, 3, 105000.00, '2024-01-01', '2024-01-31', 450, 1200, 600, 170, 25, 350, 200, 9200, '2024-02-05', 'Completed', 0, 0),   
(4, 4, 70000.00, '2024-01-01', '2024-01-31', 200, 500, 400, 160, 40, 300, 220, 7000, '2024-02-05', 'Completed', 0, 0),   
(5, 5, 115000.00, '2024-01-01', '2024-01-31', 600, 2500, 1200, 165, 35, 500, 280, 11500, '2024-02-05', 'Completed', 0, 0),   
(6, 6, 85000.00, '2024-01-01', '2024-01-31', 400, 1800, 700, 150, 30, 450, 240, 8500, '2024-02-05', 'Completed', 0, 0),   
(7, 7, 125000.00, '2024-01-01', '2024-01-31', 350, 1600, 650, 155, 28, 440, 230, 8200, '2024-02-05', 'Completed', 0, 0),   
(8, 8, 75000.00, '2024-01-01', '2024-01-31', 250, 1300, 500, 145, 32, 400, 210, 7400, '2024-02-05', 'Completed', 0, 0),   
(9, 9, 130000.00, '2024-01-01', '2024-01-31', 550, 2200, 1100, 160, 30, 500, 290, 10300, '2024-02-05', 'Completed', 0, 0),   
(10, 10, 95000.00, '2024-01-01', '2024-01-31', 500, 2100, 900, 170, 30, 480, 270, 9800, '2024-02-05', 'Completed', 0, 0),   
(11, 11, 80000.00, '2024-01-01', '2024-01-31', 450, 1800, 850, 165, 29, 470, 260, 9200, '2024-02-05', 'Completed', 0, 0),   
(12, 12, 95000.00, '2024-01-01', '2024-01-31', 500, 2200, 950, 175, 31, 480, 280, 9900, '2024-02-05', 'Completed', 0, 0),   
(13, 13, 70000.00, '2024-01-01', '2024-01-31', 400, 1700, 800, 160, 33, 460, 250, 8800, '2024-02-05', 'Completed', 0, 0),   
(14, 14, 110000.00, '2024-01-01', '2024-01-31', 350, 1600, 700, 155, 30, 440, 240, 8200, '2024-02-05', 'Completed', 0, 0),   
(15, 15, 78000.00, '2024-01-01', '2024-01-31', 250, 1400, 600, 150, 32, 420, 210, 7400, '2024-02-05', 'Completed', 0, 0),   
(16, 16, 105000.00, '2024-01-01', '2024-01-31', 300, 1500, 750, 160, 28, 430, 220, 7700, '2024-02-05', 'Completed', 0, 0),   
(17, 17, 80000.00, '2024-01-01', '2024-01-31', 450, 2000, 950, 170, 35, 490, 280, 9700, '2024-02-05', 'Completed', 0, 0),   
(18, 18, 105000.00, '2024-01-01', '2024-01-31', 400, 1800, 850, 165, 30, 460, 270, 9100, '2024-02-05', 'Completed', 0, 0),   
(19, 19, 92000.00, '2024-01-01', '2024-01-31', 300, 1500, 700, 155, 28, 430, 240, 7600, '2024-02-05', 'Completed', 0, 0),   
(20, 20, 115000.00, '2024-01-01', '2024-01-31', 500, 2200, 1000, 160, 35, 480, 290, 9900, '2024-02-05', 'Completed', 0, 0),   
(21, 21, 85000.00, '2024-01-01', '2024-01-31', 600, 2500, 1200, 175, 38, 510, 300, 11200, '2024-02-05', 'Completed', 0, 0),   
(22, 22, 100000.00, '2024-01-01', '2024-01-31', 450, 1900, 800, 165, 32, 460, 270, 9000, '2024-02-05', 'Completed', 0, 0),   
(23, 23, 80000.00, '2024-01-01', '2024-01-31', 550, 2100, 950, 160, 30, 480, 280, 9500, '2024-02-05', 'Completed', 0, 0),   
(24, 24, 120000.00, '2024-01-01', '2024-01-31', 400, 1700, 800, 155, 34, 450, 250, 8800, '2024-02-05', 'Completed', 0, 0),   
(25, 25, 95000.00, '2024-01-01', '2024-01-31', 350, 1600, 700, 150, 29, 430, 240, 8200, '2024-02-05', 'Completed', 0, 0),   
(26, 26, 105000.00, '2024-01-01', '2024-01-31', 500, 2200, 1000, 170, 32, 490, 280, 9800, '2024-02-05', 'Completed', 0, 0),   
(27, 27, 70000.00, '2024-01-01', '2024-01-31', 550, 2300, 1100, 175, 31, 500, 290, 10500, '2024-02-05', 'Completed', 0, 0),   
(28, 28, 90000.00, '2024-01-01', '2024-01-31', 400, 1800, 850, 160, 33, 470, 260, 9100, '2024-02-05', 'Completed', 0, 0),   
(29, 29, 105000.00, '2024-01-01', '2024-01-31', 300, 1500, 700, 155, 29, 430, 240, 7700, '2024-02-05', 'Completed', 0, 0),   
(30, 30, 80000.00, '2024-01-01', '2024-01-31', 450, 2000, 950, 165, 32, 490, 270, 9600, '2024-02-05', 'Completed', 0, 0),   
(31, 31, 100000.00, '2024-01-01', '2024-01-31', 500, 2100, 900, 170, 30, 480, 260, 9400, '2024-02-05', 'Completed', 0, 0),   
(32, 32, 95000.00, '2024-01-01', '2024-01-31', 400, 1800, 850, 160, 28, 460, 250, 8900, '2024-02-05', 'Completed', 0, 0),   
(33, 33, 70000.00, '2024-01-01', '2024-01-31', 250, 1300, 600, 145, 34, 420, 210, 7500, '2024-02-05', 'Completed', 0, 0),   
(34, 34, 70000.00, '2024-01-01', '2024-01-31', 600, 2400, 1200, 175, 36, 520, 310, 11200, '2024-02-05', 'Completed', 0, 0),   
(35, 35, 90000.00, '2024-01-01', '2024-01-31', 450, 1900, 850, 160, 33, 460, 270, 9000, '2024-02-05', 'Completed', 0, 0),   
(36, 36, 105000.00, '2024-01-01', '2024-01-31', 500, 2200, 1000, 170, 30, 480, 280, 9800, '2024-02-05', 'Completed', 0, 0),   
(37, 37, 80000.00, '2024-01-01', '2024-01-31', 600, 2400, 1200, 175, 35, 510, 300, 11000, '2024-02-05', 'Completed', 0, 0),   
(38, 38, 120000.00, '2024-01-01', '2024-01-31', 400, 1800, 850, 160, 32, 470, 260, 9100, '2024-02-05', 'Completed', 0, 0),   
(39, 39, 80000.00, '2024-01-01', '2024-01-31', 300, 1500, 700, 155, 29, 430, 240, 7700, '2024-02-05', 'Completed', 0, 0),   
(40, 40, 95000.00, '2024-01-01', '2024-01-31', 500, 2100, 950, 170, 30, 480, 270, 9600, '2024-02-05', 'Completed', 0, 0),   
(41, 41, 105000.00, '2024-01-01', '2024-01-31', 450, 2000, 900, 165, 32, 490, 260, 9400, '2024-02-05', 'Completed', 0, 0),   
(42, 42, 85000.00, '2024-01-01', '2024-01-31', 550, 2300, 1000, 175, 35, 500, 280, 10500, '2024-02-05', 'Completed', 0, 0),   
(43, 43, 100000.00, '2024-01-01', '2024-01-31', 400, 1700, 800, 160, 33, 450, 250, 8800, '2024-02-05', 'Completed', 0, 0),   
(44, 44, 105000.00, '2024-01-01', '2024-01-31', 300, 1500, 700, 155, 30, 430, 240, 7700, '2024-02-05', 'Completed', 0, 0),   
(45, 45, 80000.00, '2024-01-01', '2024-01-31', 400, 1800, 850, 160, 31, 470, 260, 9100, '2024-02-05', 'Completed', 0, 0),   
(46, 46, 115000.00, '2024-01-01', '2024-01-31', 500, 2100, 950, 170, 32, 480, 270, 9500, '2024-02-05', 'Completed', 0, 0),   
(47, 47, 85000.00, '2024-01-01', '2024-01-31', 600, 2500, 1200, 175, 33, 520, 310, 11300, '2024-02-05', 'Completed', 0, 0),   
(48, 48, 95000.00, '2024-01-01', '2024-01-31', 350, 1600, 700, 150, 30, 430, 240, 8100, '2024-02-05', 'Completed', 0, 0),   
(49, 49, 100000.00, '2024-01-01', '2024-01-31', 450, 2000, 850, 160, 32, 470, 250, 9000, '2024-02-05', 'Completed', 0, 0),   
(50, 50, 110000.00, '2024-01-01', '2024-01-31', 500, 2100, 900, 170, 28, 490, 260, 9400, '2024-02-05', 'Completed', 0, 0);   


 -- step 6: Insert data into the reason category table   

 INSERT INTO [EmployeeDatas].[REASON_CATEGORY] (reasonCategoryID, reasonName)   

VALUES   
(1, 'Performance'),   
(2, 'Behavioral'),   
(3, 'Attendance'),   
(4, 'Other');   


--step 7: insert the offboarding data  


INSERT INTO [EmployeeDatas].[OFFBOARDING_DECISION]   

(decisionID, employeeID, reviewDate, decisionMaker, decisionStatus, terminationDate, severanceAmount, reasonCategoryID, performanceScore, replacementNeeded, exitInterviewScheduled, legalReviewComplete, appealStatus)   

VALUES   
(1, 12, '2025-03-05', 1, 'Terminated', '2025-03-10', 1500.00, 3, 45, 1, 1, 1, 'Pending'),   
(2, 22, '2025-03-06', 5, 'Terminated', '2025-03-12', 2000.00, 1, 38, 1, 1, 1, 'Pending'),   
(3, 15, NULL, NULL, NULL, NULL, NULL, NULL, 72, 0, 0, 0, NULL),   
(4, 18, NULL, NULL, NULL, NULL, NULL, NULL, 80, 0, 0, 0, NULL),   
(5, 10, NULL, NULL, NULL, NULL, NULL, NULL, 95, 0, 0, 0, NULL),   
(6, 5, NULL, NULL, NULL, NULL, NULL, NULL, 88, 0, 0, 0, NULL),   
(7, 20, NULL, NULL, NULL, NULL, NULL, NULL, 65, 1, 0, 0, NULL),   
(8, 25, NULL, NULL, NULL, NULL, NULL, NULL, 70, 1, 0, 0, NULL);   


  -- step 8: Insert data into the performance data table    

INSERT INTO [EmployeeDatas].[PERFORMANCE_DATA]   

(performanceID, employeeID, evaluationDate, utilizationRate, clientFeedbackScore, revenueGenerated, teamworkRating, performanceRating)   

VALUES   
(1, 1, '2025-01-15', 85.5, 4, 20000.00, 5, 4),   
(2, 2, '2025-01-18', 90.0, 2, 25000.00, 4, 4),   
(3, 3, '2025-01-20', 78.0, 3, 15000.00, 4, 3),   
(4, 4, '2025-01-22', 88.5, 4, 22000.00, 5, 5),   
(5, 5, '2025-01-25', 92.0, 5, 27000.00, 4, 4),   
(6, 6, '2025-01-30', 77.5, 3, 16000.00, 3, 3),   
(7, 7, '2025-02-02', 80.0, 2, 13000.00, 4, 4),   
(8, 8, '2025-02-05', 91.0, 4, 24000.00, 5, 5),   
(9, 9, '2025-02-07', 86.0, 4, 21000.00, 4, 4),   
(10, 10, '2025-02-10', 89.5, 4, 23000.00, 2, 5),   
(11, 11, '2025-02-12', 75.5, 3, 1000.00, 3, 3),   
(12, 12, '2025-02-15', 83.0, 4, 20000.00, 4, 4),   
(13, 13, '2025-02-18', 87.5, 5, 25000.00, 5, 5),   
(14, 14, '2025-02-20', 82.0, 4, 18000.00, 3, 3),   
(15, 15, '2025-02-22', 79.0, 3, 17000.00, 4, 4),   
(16, 16, '2025-02-25', 90.5, 5, 26000.00, 5, 5),   
(17, 17, '2025-02-28', 84.5, 2, 12000.00, 2, 2),   
(18, 18, '2025-03-02', 93.0, 2, 28000.00, 2, 4),   
(19, 19, '2025-03-05', 88.0, 2, 22000.00, 5, 5),   
(20, 20, '2025-03-07', 81.0, 3, 15000.00, 4, 2),   
(21, 21, '2025-03-10', 85.5, 5, 24000.00, 5, 5),   
(22, 22, '2025-03-12', 79.5, 2, 13000.00, 4, 4),   
(23, 23, '2025-03-15', 86.5, 4, 21000.00, 4, 4),   
(24, 24, '2025-03-17', 91.0, 5, 25000.00, 5, 5),   
(25, 25, '2025-03-20', 77.0, 3, 16000.00, 3, 3),   
(26, 26, '2025-03-22', 84.0, 4, 20000.00, 4, 4),   
(27, 27, '2025-03-25', 89.0, 5, 13000.00, 5, 5),   
(28, 28, '2025-03-27', 92.5, 5, 28000.00, 4, 3),   
(29, 29, '2025-03-30', 80.5, 3, 15000.00, 4, 3),   
(30, 30, '2025-04-02', 75.0, 2, 13000.00, 2, 2),   
(31, 31, '2025-04-05', 93.5, 5, 29000.00, 5, 5),   
(32, 32, '2025-04-07', 85.0, 4, 22000.00, 4, 4),   
(33, 33, '2025-04-10', 88.0, 4, 23000.00, 4, 4),   
(34, 34, '2025-04-12', 79.5, 3, 17000.00, 3, 3),   
(35, 35, '2025-04-15', 91.5, 5, 26000.00, 5, 5),   
(36, 36, '2025-04-17', 85.5, 4, 9000.00, 4, 4),   
(37, 37, '2025-04-20', 90.0, 5, 27000.00, 5, 2),   
(38, 38, '2025-04-22', 82.5, 3, 18000.00, 4, 3),   
(39, 39, '2025-04-25', 89.5, 5, 28000.00, 5, 5),   
(40, 40, '2025-04-28', 76.0, 2, 14000.00, 4, 4),   
(41, 41, '2025-05-01', 87.5, 4, 25000.00, 3, 3),   
(42, 42, '2025-05-03', 91.0, 5, 27000.00, 5, 2),   
(43, 43, '2025-05-05', 80.0, 3, 16000.00, 4, 4),   
(44, 44, '2025-05-08', 84.0, 4, 21000.00, 5, 4),   
(45, 45, '2025-05-10', 78.5, 3, 15000.00, 4, 3),   
(46, 46, '2025-05-12', 93.0, 5, 28000.00, 5, 5),   
(47, 47, '2025-05-15', 89.0, 4, 24000.00, 4, 4),   
(48, 48, '2025-05-17', 82.0, 4, 23000.00, 3, 3),   
(49, 49, '2025-05-20', 91.5, 5, 29000.00, 5, 5),   
(50, 50, '2025-05-22', 77.5, 3, 17000.00, 4, 3);   

 --step 9 insert skill compendium:  

INSERT INTO [EmployeeDatas].SKILL_COMPENDIUM (skill_name, skill_category, skill_description)   

VALUES   
('Python Programming', 'Programming', 'Proficiency in Python for data analysis and web development.'),   
('SQL Database Management', 'Database', 'Expertise in writing complex SQL queries and database administration.'),   
('Project Management', 'Management', 'Ability to plan, execute, and close projects effectively.'),   
('Data Analysis', 'Data Science', 'Skills in analyzing large datasets to derive meaningful insights.'),   
('Customer Service', 'Soft Skills', 'Excellent communication and problem-solving skills for customer interactions.'),   
('Java Development', 'Programming', 'Experience in developing Java applications for enterprise solutions.'),   
('Network Administration', 'IT', 'Knowledge of network infrastructure and troubleshooting.'),   
('Financial Modeling', 'Finance', 'Creating and analyzing financial models for forecasting and valuation.'),   
('Digital Marketing', 'Marketing', 'Developing and executing digital marketing campaigns.'),   
('Technical Writing', 'Communication', 'Creating clear and concise technical documentation.'),   
('Cloud Computing (AWS)', 'IT', 'Experience with Amazon Web Services for cloud-based solutions.'),   
('UI/UX Design', 'Design', 'Designing user-friendly interfaces and experiences.'),   
('Statistical Analysis', 'Data Science', 'Applying statistical methods to analyze data and draw conclusions.'),   
('Human Resources Management', 'HR', 'Managing employee relations and HR processes.'),   
('Sales Negotiation', 'Sales', 'Negotiating and closing sales deals effectively.'),   
('Machine Learning', 'Data Science', 'Implementing machine learning algorithms for predictive analysis.'),   
('Cybersecurity', 'IT', 'Protecting systems and data from cyber threats.'),   
('Content Creation', 'Marketing', 'Creating engaging content for various platforms.'),   
('Public Speaking', 'Communication', 'Delivering effective presentations and speeches.'),   
('Agile Methodology', 'Management', 'Applying agile principles and practices to project development.');   
 

  --step 10; insert employee specific skills  

INSERT INTO [EmployeeDatas].EMPLOYEE_SKILLS (EMPLOYEE_SKILLS_ID, EMPLOYEE_ID, SKILL_ID, PROFICIENCY_LEVEL, LAST_USED)   

VALUES    
(1, 1, 1, 'Expert', '2024-10-26'),  -- Employee 1: Python (Expert)   
(2, 1, 2, 'Advanced', '2024-10-20'), -- Employee 1: SQL (Advanced)   
(3, 2, 3, 'Intermediate', '2024-10-25'), -- Employee 2: Project Management (Intermediate)   
(4, 2, 14, 'Beginner', '2024-10-15'), -- Employee 2: Data Analysis (Beginner)   
(5, 3, 5, 'Advanced', '2024-10-27'), -- Employee 3: Customer Service (Advanced)   
(6, 3, 6, 'Intermediate', '2024-10-22'), -- Employee 3: Java Development (Intermediate)   
(7, 4, 17, 'Expert', '2024-10-28'), -- Employee 4: Network Administration (Expert)   
(8, 4, 8, 'Advanced', '2024-10-23'), -- Employee 4: Financial Modeling (Advanced)   
(9, 5, 20, 'Intermediate', '2024-10-26'), -- Employee 5: Digital Marketing (Intermediate)   
(10, 5, 10, 'Beginner', '2024-10-18'), -- Employee 5: Technical Writing (Beginner)   
(11, 6, 13, 'Advanced', '2024-10-27'),  -- Employee 6: Python (Advanced)   
(12, 6, 11, 'Intermediate', '2024-10-21'), -- Employee 6: Cloud Computing (AWS) (Intermediate)   
(13, 7, 5, 'Expert', '2024-10-29'), -- Employee 7: UI/UX Design (Expert)   
(14, 7, 13, 'Advanced', '2024-10-24'), -- Employee 7: Statistical Analysis (Advanced)   
(15, 8, 14, 'Intermediate', '2024-10-27'), -- Employee 8: HR Management (Intermediate)   
(16, 8, 13, 'Beginner', '2024-10-19'), -- Employee 8: Sales Negotiation (Beginner)   
(17, 9, 16, 'Advanced', '2024-10-28'), -- Employee 9: Machine Learning (Advanced)   
(18, 9, 17, 'Intermediate', '2024-10-22'), -- Employee 9: Cybersecurity (Intermediate)   
(19, 10, 12, 'Expert', '2024-10-30'), -- Employee 10: Content Creation (Expert)   
(20, 10, 19, 'Advanced', '2024-10-25'), -- Employee 10: Public Speaking (Advanced)   
(21, 11, 20, 'Intermediate', '2024-10-28'), -- Employee 11: Agile Methodology (Intermediate)   
(22, 11, 12, 'Beginner', '2024-10-20'), -- Employee 11: SQL (Beginner)   
(23, 12, 3, 'Advanced', '2024-10-29'),  -- Employee 12: Project Management (Advanced)   
(24, 12, 4, 'Intermediate', '2024-10-23'), -- Employee 12: Data Analysis (Intermediate)   
(25, 13, 20, 'Expert', '2024-10-30'),  -- Employee 13: Customer Service (Expert)   
(26, 13, 6, 'Advanced', '2024-10-24'), -- Employee 13: Java Development (Advanced)   
(27, 14, 6, 'Intermediate', '2024-10-27'), -- Employee 14: Network Administration (Intermediate)   
(28, 14, 8, 'Beginner', '2024-10-19'), -- Employee 14: Financial Modeling (Beginner)   
(29, 15, 13, 'Advanced', '2024-10-28'),  -- Employee 15: Digital Marketing (Advanced)   
(30, 15, 10, 'Intermediate', '2024-10-22'), -- Employee 15: Technical Writing (Intermediate)   
(31, 16, 20, 'Expert', '2024-10-31'), -- Employee 16: Cloud Computing (AWS) (Expert)   
(32, 16, 12, 'Advanced', '2024-10-25'), -- Employee 16: UI/UX Design (Advanced)   
(33, 17, 11, 'Intermediate', '2024-10-28'), -- Employee 17: Statistical Analysis (Intermediate)   
(34, 17, 14, 'Beginner', '2024-10-20'), -- Employee 17: HR Management (Beginner)   
(35, 18, 15, 'Advanced', '2024-10-29'),  -- Employee 18: Sales Negotiation (Advanced)   
(36, 18, 20, 'Intermediate', '2024-10-23'), -- Employee 18: Machine Learning (Intermediate)   
(37, 19, 17, 'Expert', '2024-10-30'),  -- Employee 19: Cybersecurity (Expert)   
(38, 19, 16, 'Advanced', '2024-10-24'), -- Employee 19: Content Creation (Advanced)   
(39, 20, 19, 'Intermediate', '2024-10-27'), -- Employee 20: Public Speaking (Intermediate)   
(40, 20, 20, 'Beginner', '2024-10-19'), -- Employee 20: Agile Methodology (Beginner)   
(41, 21, 1, 'Advanced', '2024-10-31'),  -- Employee 21: Python Programming (Advanced)   
(42, 21, 2, 'Intermediate', '2024-10-25'), -- Employee 21: SQL Database Management (Intermediate)   
(43, 22, 3, 'Expert', '2024-11-01'),  -- Employee 22: Project Management (Expert)   
(44, 22, 4, 'Advanced', '2024-10-26'), -- Employee 22: Data Analysis (Advanced)   
(45, 23, 6, 'Intermediate', '2024-10-29'), -- Employee 23: Customer Service (Intermediate)   
(46, 23, 1, 'Beginner', '2024-10-21'), -- Employee 23: Java Development (Beginner)   
(47, 24, 7, 'Advanced', '2024-10-30'),  -- Employee 24: Network Administration (Advanced)   
(48, 24, 8, 'Intermediate', '2024-10-24'), -- Employee 24: Financial Modeling (Intermediate)   
(49, 25, 12, 'Expert', '2024-11-02'),  -- Employee 25: Digital Marketing (Expert)   
(50, 25, 10, 'Advanced', '2024-10-27'), -- Employee 25: Technical Writing (Advanced)   
(51, 26, 11, 'Intermediate', '2024-10-30'), -- Employee 26: Cloud Computing (AWS) (Intermediate)   
(52, 26, 12, 'Beginner', '2024-10-22'), -- Employee 26: UI/UX Design (Beginner)   
(53, 27, 7, 'Advanced', '2024-10-31'),  -- Employee 27: Statistical Analysis (Advanced)   
(54, 27, 20, 'Intermediate', '2024-10-25'), -- Employee 27: Human Resources Management (Intermediate)   
(55, 28, 15, 'Expert', '2024-11-03'),  -- Employee 28: Sales Negotiation (Expert)   
(56, 28, 12, 'Advanced', '2024-10-28'), -- Employee 28: Machine Learning (Advanced)   
(57, 29, 17, 'Intermediate', '2024-10-31'), -- Employee 29: Cybersecurity (Intermediate)   
(58, 29, 20, 'Beginner', '2024-10-23'), -- Employee 29: Content Creation (Beginner)   
(59, 30, 19, 'Advanced', '2024-11-01'),  -- Employee 30: Public Speaking (Advanced)   
(60, 30, 20, 'Intermediate', '2024-10-26'), -- Employee 30: Agile Methodology (Intermediate)   
(61, 31, 20, 'Expert', '2024-11-04'),    -- Employee 31: Python Programming (Expert)   
(62, 31, 2, 'Advanced', '2024-10-29'),   -- Employee 31: SQL Database Management (Advanced)   
(63, 32, 12, 'Intermediate', '2024-11-02'), -- Employee 32: Project Management (Intermediate)   
(64, 32, 4, 'Beginner', '2024-10-24'),    -- Employee 32: Data Analysis (Beginner)   
(65, 33, 5, 'Advanced', '2024-11-03'),   -- Employee 33: Customer Service (Advanced)   
(66, 33, 6, 'Intermediate', '2024-10-27'), -- Employee 33: Java Development (Intermediate)   
(67, 34, 7, 'Expert', '2024-11-05'),    -- Employee 34: Network Administration (Expert)   
(68, 34, 16, 'Advanced', '2024-10-30'),   -- Employee 34: Financial Modeling (Advanced)   
(69, 35, 9, 'Intermediate', '2024-11-03'), -- Employee 35: Digital Marketing (Intermediate)   
(70, 35, 10, 'Beginner', '2024-10-25'),   -- Employee 35: Technical Writing (Beginner)   
(71, 36, 2, 'Advanced', '2024-11-04'),   -- Employee 36: Cloud Computing (AWS) (Advanced)   
(72, 36, 12, 'Intermediate', '2024-10-28'), -- Employee 36: UI/UX Design (Intermediate)   
(73, 37, 3, 'Expert', '2024-11-06'),    -- Employee 37: Statistical Analysis (Expert)   
(74, 37, 14, 'Advanced', '2024-10-31'),   -- Employee 37: Human Resources Management (Advanced)   
(75, 38, 15, 'Intermediate', '2024-11-04'), -- Employee 38: Sales Negotiation (Intermediate)   
(76, 38, 16, 'Beginner', '2024-10-26'),   -- Employee 38: Machine Learning (Beginner)   
(77, 39, 17, 'Advanced', '2024-11-05'),   -- Employee 39: Cybersecurity (Advanced)   
(78, 39, 20, 'Intermediate', '2024-10-29'), -- Employee 39: Content Creation (Intermediate)   
(79, 40, 19, 'Expert', '2024-11-07'),    -- Employee 40: Public Speaking (Expert)   
(80, 40, 20, 'Advanced', '2024-11-01'),   -- Employee 40: Agile Methodology (Advanced)   
(81, 41, 1, 'Intermediate', '2024-11-06'), -- Employee 41: Python Programming (Intermediate)   
(82, 41, 2, 'Beginner', '2024-10-31'),    -- Employee 41: SQL Database Management (Beginner)   
(83, 42, 3, 'Advanced', '2024-11-04'),   -- Employee 42: Project Management (Advanced)   
(84, 42, 2, 'Intermediate', '2024-10-29'), -- Employee 42: Data Analysis (Intermediate)   
(85, 43, 5, 'Expert', '2024-11-07'),    -- Employee 43: Customer Service (Expert)   
(86, 43, 6, 'Advanced', '2024-11-02'),   -- Employee 43: Java Development (Advanced)   
(87, 44, 7, 'Intermediate', '2024-11-05'), -- Employee 44: Network Administration (Intermediate)   
(88, 44, 8, 'Beginner', '2024-10-30'),    -- Employee 44: Financial Modeling (Beginner)   
(89, 45, 9, 'Advanced', '2024-11-08'),   -- Employee 45: Digital Marketing (Advanced)   
(90, 45, 10, 'Intermediate', '2024-11-03'), -- Employee 45: Technical Writing (Intermediate)   
(91, 46, 11, 'Expert', '2024-11-07'),    -- Employee 46: Cloud Computing (AWS) (Expert)   
(92, 46, 9, 'Advanced', '2024-11-01'),   -- Employee 46: UI/UX Design (Advanced)   
(93, 47, 13, 'Intermediate', '2024-11-06'), -- Employee 47: Statistical Analysis (Intermediate)   
(94, 47, 3, 'Beginner', '2024-10-31'),    -- Employee 47: Human Resources Management (Beginner)   
(95, 48, 18, 'Advanced', '2024-11-09'),   -- Employee 48: Sales Negotiation (Advanced)   
(96, 48, 1, 'Intermediate', '2024-11-04'), -- Employee 48: Machine Learning (Intermediate)   
(97, 49, 17, 'Expert', '2024-11-08'),    -- Employee 49: Cybersecurity (Expert)   
(98, 49, 6, 'Advanced', '2024-11-03'),   -- Employee 49: Content Creation (Advanced)   
(99, 50, 2, 'Intermediate', '2024-11-06'), -- Employee 50: Public Speaking (Intermediate)   
(100, 50, 1, 'Beginner', '2024-10-31');  -- Employee 50: Agile Methodology (Beginner)   


-- step 11: add the employee table foreign keys   


ALTER TABLE [EmployeeDatas].[EMPLOYEE]  
ADD CONSTRAINT FK_EMP_DEPT  
FOREIGN KEY (departmentID)  
REFERENCES [EmployeeDatas].[DEPARTMENT](departmentID);  
 
ALTER TABLE [EmployeeDatas].[EMPLOYEE]  
ADD CONSTRAINT FK_EMP_MENTOR  
FOREIGN KEY (mentorID)  
REFERENCES [EmployeeDatas].[EMPLOYEE](employeeID)  

ON DELETE NO ACTION;  


ALTER TABLE [EmployeeDatas].[EMPLOYEE]  
ADD CONSTRAINT FK_EMP_PAYROLL  
FOREIGN KEY (payrollID)  
REFERENCES [EmployeeDatas].[PAYROLL](payrollID);  
 

ALTER TABLE [EmployeeDatas].[EMPLOYEE]  
ADD CONSTRAINT FK_EMP_PERFORMANCE  
FOREIGN KEY (performanceID)  
REFERENCES [EmployeeDatas].[PERFORMANCE_DATA](performanceID); 

 
  



-- Queries

SELECT e.employeeID, e.firstName, e.lastName, pd.performanceRating, pd.utilizationRate 
FROM [EmployeeDatas].[EMPLOYEE] e 
JOIN [EmployeeDatas].[PERFORMANCE_DATA] pd ON e.employeeID = pd.employeeID 
WHERE pd.performanceRating <= 3 OR pd.utilizationRate <=80
ORDER BY pd.performanceRating ASC; 


SELECT e.employeeID, e.firstName, e.lastName, p.nonBillableHours 
FROM [EmployeeDatas].[EMPLOYEE] e 
JOIN [EmployeeDatas].[PAYROLL] p ON e.employeeID = p.employeeID 
WHERE p.nonBillableHours > 34; 


SELECT e.employeeID, e.firstName, e.lastName, od.appealStatus, od.reviewDate 
FROM [EmployeeDatas].[EMPLOYEE] e 
JOIN [EmployeeDatas].[OFFBOARDING_DECISION] od ON e.employeeID = od.employeeID 
WHERE od.appealStatus IN ('Pending');


SELECT e.employeeID, e.firstName, e.lastName, pd.utilizationRate, pd.revenueGenerated 
FROM [EmployeeDatas].[EMPLOYEE] e 
JOIN [EmployeeDatas].[PERFORMANCE_DATA] pd ON e.employeeID = pd.employeeID 
ORDER BY pd.revenueGenerated ASC 
