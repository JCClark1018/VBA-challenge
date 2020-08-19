----- 1. List the following details of each employee: employee number, last name, first name, sex, and salary.
--Select * From employee_info
--Select * From salaries

Select ei.emp_no,
	ei.last_name,
	ei.first_name,
	ei.sex,
	sa.salary
From employee_info as ei
	Inner Join salaries as sa
		On (ei.emp_no = sa.emp_no)
Order By ei.emp_no;

----- 2. List first name, last name, and hire date for employees who were hired in 1986.
Select first_name,
	last_name,
	hire_date
From employee_info
Where hire_date Between '1986-01-01' And '1986-12-31'

----- 3. List the manager of each department with the following information: department number, department name, 
----- the manager's employee number, last name, first name.

Select
	di.dept_no,
	di.dept_name,
	dm.emp_no,
	ei.last_name,
	ei.first_name
	
From department_info as di
	Inner Join department_manager as dm
		On (di.dept_no = dm.dept_no)
	Inner Join employee_info as ei
		On (dm.emp_no = ei.emp_no)
		
Order By di.dept_no;		

----- 4. List the department of each employee with the following information: employee number, last name, first name, and department name.	

Select ei.emp_no,
	ei.last_name,
	ei.first_name,
	di.dept_name
From employee_info as ei
	Inner Join employee_dept as ed
		On (ei.emp_no = ed.emp_no)
	Inner Join department_info as di 
		On (ed.dept_no = di.dept_no)
Order By ei.emp_no

----- 5. List first name, last name, and sex for employees whose first name is "Hercules" and last names begin with "B."
Select ei.first_name,
	ei.last_name,
	ei.sex
From employee_info as ei
Where ei.first_name = 'Hercules' and ei.last_name like 'B%'

----- 6. List all employees in the Sales department, including their employee number, last name, first name, and department name.
Select ei.emp_no,
	ei.last_name,
	ei.first_name,
	di.dept_name
From employee_info as ei
	Inner Join employee_dept as ed 
		On (ei.emp_no = ed.emp_no)
	Inner Join department_info as di
		On (ed.dept_no = di.dept_no)
Where di.dept_name = 'Sales'
		
----- 7. List all employees in the Sales and Development departments, including their employee number, last name, first name, and department name.
Select ei.emp_no,
	ei.last_name,
	ei.first_name,
	di.dept_name
From employee_info as ei
	Inner Join employee_dept as ed 
		On (ei.emp_no = ed.emp_no)
	Inner Join department_info as di
		On (ed.dept_no = di.dept_no)
Where di.dept_name = 'Sales' OR di.dept_name = 'Development'

----- 8. In descending order, list the frequency count of employee last names, i.e., how many employees share each last name.
Select
	last_name,
	Count (last_name)
From
	employee_info
Group By
	last_name
Order By Count Desc;
