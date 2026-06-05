
from admission import add_student
from search import search_student


while True:

    print("\n===== SCHOOL MANAGEMENT SYSTEM =====")

    print("1. New Admission")
    print("2. Search Student")
    print("3. Exit")

    choice = input("\nEnter Your Choice: ")

    if choice == "1":
        add_student()

    elif choice == "2":
        search_student()

    elif choice == "3":
        print("\nProgram Closed")
        break

    else:
        print("\nInvalid Choice")

