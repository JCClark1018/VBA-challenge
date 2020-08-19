Create Table title_ids (
	title_id VARCHAR(50) Not Null,
	title VARCHAR(50) Not Null,
	Primary Key (title_id) 
);

Create Table employee_info (
	emp_no INT Not NULL,
	emp_title_id VARCHAR(50) Not Null,
	birth_date Date Not NULL,
	first_name VARCHAR(50) Not NULL,
	last_name VARCHAR(50) Not NULL,
	sex VARCHAR(50) Not NULL,
	hire_date DATE Not NULL,
	Foreign Key (emp_title_id) References title_ids (title_id),
	Primary Key (emp_no)
);

Create Table department_info (
	dept_no VARCHAR(50) Not NULL,
	dept_name VARCHAR(50) Not NULL,
	Primary Key (dept_no),
	Unique(dept_name)
);

Create Table department_manager (
	dept_no VARCHAR(50) Not NULL,
	emp_no INT Not NULL,
	Foreign Key (emp_no) References employee_info (emp_no),
	Foreign Key (dept_no) References department_info (dept_no),
	Primary Key (emp_no, dept_no)
);

Create Table employee_dept (
	emp_no INT Not Null,
	dept_no VARCHAR(50) Not Null,
	Foreign Key (emp_no) References employee_info (emp_no),
	Foreign Key (dept_no) References department_info (dept_no),
	Primary Key (emp_no, dept_no)
);

Create Table salaries (
	emp_no INT Not Null,
	salary INT Not Null,
	Foreign Key (emp_no) References employee_info (emp_no),
	Primary Key (emp_no, salary)
);


