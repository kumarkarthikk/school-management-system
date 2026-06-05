
from openpyxl import load_workbook


# Open Excel file
wb = load_workbook(r"C:\Users\burle\OneDrive\Desktop\py_projects\1st_project\students.xlsx")

sheet = wb.active


students = [

    [1, "01-06-2026", "Rahul Kumar", 10, "Ramesh Kumar", "Sita Devi", "Chennai", "9876543210", "ABC Public School", "ACTIVE", "New Admission"],

    [2, "01-06-2026", "Priya Sharma", 9, "Mohan Sharma", "Kavitha", "Hyderabad", "9123456780", "Sunrise School", "ACTIVE", "Scholarship Student"],

    [3, "02-06-2026", "Arjun Reddy", 8, "Ravi Reddy", "Lakshmi", "Vijayawada", "9012345678", "Sri Chaitanya", "ACTIVE", "Good in Sports"],

    [4, "02-06-2026", "Sneha Patel", 7, "Mahesh Patel", "Rekha Patel", "Mumbai", "9988776655", "Oxford School", "ACTIVE", "Transfer Student"],

    [5, "03-06-2026", "Kiran Kumar", 6, "Suresh Kumar", "Anitha", "Bangalore", "9090909090", "Delhi Public School", "ACTIVE", "Hostel Student"],

    [6, "03-06-2026", "Meena Das", 5, "Raju Das", "Sunitha", "Kolkata", "9345678910", "Little Flower", "ACTIVE", "New Admission"],

    [7, "04-06-2026", "Vikram Singh", 10, "Ajay Singh", "Pooja Singh", "Delhi", "9871234560", "St Xavier School", "ACTIVE", "Excellent Marks"],

    [8, "04-06-2026", "Anjali Rao", 8, "Ganesh Rao", "Bhavani Rao", "Pune", "9765432109", "Narayana School", "ACTIVE", "NCC Student"],

    [9, "05-06-2026", "Rohit Verma", 9, "Ashok Verma", "Latha Verma", "Visakhapatnam", "9556677889", "Vidya Niketan", "ACTIVE", "Fee Pending"],

    [10, "05-06-2026", "Divya Nair", 7, "Suresh Nair", "Deepa Nair", "Kochi", "9445566778", "Kendriya Vidyalaya", "ACTIVE", "Transfer Student"],

    [11, "06-06-2026", "Naveen Kumar", 6, "Hari Kumar", "Jyothi", "Tirupati", "9332211445", "Bhashyam School", "ACTIVE", "Bus Facility"],

    [12, "06-06-2026", "Pooja Mehta", 5, "Raj Mehta", "Shalini Mehta", "Ahmedabad", "9223344556", "DAV School", "ACTIVE", "New Admission"],

    [13, "07-06-2026", "Akash Yadav", 8, "Vinod Yadav", "Uma Yadav", "Patna", "9112233445", "Green Valley School", "ACTIVE", "Good Discipline"],

    [14, "07-06-2026", "Neha Kapoor", 9, "Rakesh Kapoor", "Renu Kapoor", "Chandigarh", "9001122334", "Modern School", "ACTIVE", "Sports Quota"],

    [15, "08-06-2026", "Sai Teja", 10, "Krishna Rao", "Lalitha Rao", "Guntur", "9898989898", "Bhashyam School", "ACTIVE", "Topper"]

]


# Add rows into Excel
for student in students:

    sheet.append(student)


# Save Excel file
wb.save(r"C:\Users\burle\OneDrive\Desktop\py_projects\1st_project\students.xlsx")


print("15 Dummy Records Added Successfully")

