import time
average = 0
student_list = ["Alfred", "Matthew", "Simon", "Amy", "Mary"]
Selection = False
Average = False
scores = {
    "Alfred": {
        "Maths": 74,
        "Science": 78,
        "English": 68,
        "History": 52
    },
    "Matthew": {
        "Maths": 64,
        "Science": 82,
        "English": 70,
        "History": 65,
        "Psychology": 87
    },
    "Simon": {
        "Maths": 61,
        "Science": 63,
        "English": 71,
        "History": 56
    },
    "Amy": {
        "Maths": 91,
        "Science": 89,
        "English": 69,
        "History": 71
    },
    "Mary": {
        "Maths": 53,
        "Science": 74,
        "English": 90,
        "History": 85,
        "Psychology": 75
    },
}

print("Hi, welcome to the Student Averagator v1.")
time.sleep(1)
while Selection is False:
    print("Please choose from the following students: \n")
    time.sleep(1)
    print("Alfred, \n"
      "Amy, \n"
      "Mary, \n"
      "Matthew, \n"
      "Simon \n")
    Student = input("SELECT YOUR STUDENT: \n").capitalize()
    if Student in student_list:
        Selection = True
    else:
        print("Please type in a current student.")
while Selection is True:
    for name in scores:
        print(Student + "scores:")
        print(scores[Student])
        Selection = False


while average is False:
    average = input("Would you like the average:\n")
    if average == ["y", "yes"]:
        for Student in scores:
            print(scores[Student])
            for subject in scores:
                for marks in scores:
                    print(sum(marks) / 5)
    average = True

    if average == ["n", "no"]:
        print("Thanks for using Student Averagator. Goodbye")
        time.sleep(1)
        exit()
    else:
        print("Please type yes or no.")

while average is True:
    print("Thanks for using Student Averagtor. Goodbye")