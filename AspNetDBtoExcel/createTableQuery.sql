CREATE TABLE [dbo].[EmployeeInfo] (
    [EmpNo]       INT          IDENTITY (1, 1) NOT NULL,
    [EmpName]     VARCHAR (50) NOT NULL,
    [Salary]      INT          NOT NULL,
    [DeptName]    VARCHAR (50) NOT NULL,
    [Designation] VARCHAR (50) NOT NULL,
    [HRA]         AS           ([Salary]*(0.2)),
    [TA]          AS           ([Salary]*(0.15)),
    [DA]          AS           ([Salary]*(0.18)),
    [GrossSalary] AS           ((([Salary]+[Salary]*(0.2))+[Salary]*(0.15))+[Salary]*(0.18)),
    [TDS]         AS           (((([Salary]+[Salary]*(0.2))+[Salary]*(0.15))+[Salary]*(0.18))*(0.25)),
    [NetSalary]   AS           (((([Salary]+[Salary]*(0.2))+[Salary]*(0.15))+[Salary]*(0.18))-((([Salary]+[Salary]*(0.2))+[Salary]*(0.15))+[Salary]*(0.18))*(0.25)),
    PRIMARY KEY CLUSTERED ([EmpNo] ASC)
);