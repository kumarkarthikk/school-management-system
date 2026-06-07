
from admission import add_student
from search import search_student
from update import update_student
from leave import leave_student


while True:

    print("\n===== SCHOOL MANAGEMENT SYSTEM =====")

    print("1. New Admission")
    print("2. Search Student")
    print("3. Update Student")
    print("4. Leave Student")
    print("5. Exit")

    choice = input("\nEnter Your Choice: ")

    if choice == "1":
        add_student()

    elif choice == "2":
        search_student()

    elif choice == "3":
        update_student()

    elif choice == "4":
        leave_student()

    elif choice == "5":
        print("\nProgram Closed")
        break

    else:
        print("\nInvalid Choice")

