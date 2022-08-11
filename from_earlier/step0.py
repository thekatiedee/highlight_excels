import sys

# exit if user says the files don't match
def exit():
    print("exiting now...")
    sys.exit()


def intro():

    print("")
    print(
        "to run this program, i will need the following documents: sheet1, fin ops, IMG sheet"
    )
    print("")
    print(
        "your documents need to be in this file format blah blah blah and go into excel_docs folder yada yada"
    )

    answer = input("have you updated all your files? y/n")
    answer_low = answer.lower()
    while True:
        if answer_low == "y":
            break
        elif answer_low == "n":
            exit()
        else:
            answer = input(
                "you did not enter an answer i can understand. did you save the right files? y/n"
            )
            answer_low = answer.lower()


if __name__ == "__main__":
    intro()
