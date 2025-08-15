import win32com.client as wincl

speaker_number = 0
speaker = wincl.Dispatch("SAPI.SpVoice")
voices = speaker.GetVoices()

speaker.Speak(f"Hello, my name is {voices.Item(speaker_number).GetAttribute('Name')}")
print("\t***** TYPE OF GREETINGS *****\n")
print("\t1.Good Morning, ")
print("\t2.Good Afternoon, ")
print("\t3.Good Evening, ")
print("\t4.Good Night, ")
speaker.Speak("click 1 for good Morning wishes ,click 2 for good afternoon wishes,click 3 for good evening wishes,click 4 for good night wishes")
wish=int(input("Enter the choice :"))
if(wish==1):
    speaker.Speak(f"Your Choice was Good Morning ")
elif(wish==2):
    speaker.Speak(f"Your Choice was Good Afternoon ")
elif(wish==3):
    speaker.Speak(f"Your Choice was Good evening ")
else:
    speaker.Speak(f"Your Choice was good night ")
match wish:
    case 1:
        print("Enter the names for whom you want to greet a good morning message")
        speaker.Speak("Enter the names for whom you want to greet a good morning message")
        name_list = []
        n = int(input("Enter the number of people to greet: "))
        speaker.Speak("Enter the number of people to greet: ")

        for i in range(n):
            speaker.Speak("Enter the name: ")
            names = input("Enter the name: ")
            
            name_list.append(names)

        for name in name_list:
            speaker.Speak(f"Good morning, {name}! , Have a nice day")
    case 2:
        print("Enter the names for whom you want to greet a Good Afternoon message")
        speaker.Speak("Enter the names for whom you want to greet a Good Afternoon message")
        name_list = []
        n = int(input("Enter the number of people to greet: "))
        speaker.Speak("Enter the number of people to greet: ")

        for i in range(n):
            speaker.Speak("Enter the name: ")
            names = input("Enter the name: ")
            
            name_list.append(names)

        for name in name_list:
            speaker.Speak(f"Good Afternoon, {name}!")
    case 3:
        print("Enter the names for whom you want to greet a good evening message")
        speaker.Speak("Enter the names for whom you want to greet a good evening message")
        name_list = []
        n = int(input("Enter the number of people to greet: "))
        speaker.Speak("Enter the number of people to greet: ")

        for i in range(n):
            speaker.Speak("Enter the name: ")
            names = input("Enter the name: ")
            
            name_list.append(names)

        for name in name_list:
            speaker.Speak(f"Good evening, {name}! , how was your day")
    case 4:
        print("Enter the names for whom you want to greet a good night message")
        speaker.Speak("Enter the names for whom you want to greet a good night message")
        name_list = []
        n = int(input("Enter the number of people to greet: "))
        speaker.Speak("Enter the number of people to greet: ")

        for i in range(n):
            speaker.Speak("Enter the name: ")
            names = input("Enter the name: ")
            
            name_list.append(names)

        for name in name_list:
            speaker.Speak(f"Good night, {name}! , sleep well")
    case _:
        print("you have entered a wrong Choice Please try again")
        speaker.Speak("you have entered a wrong Choice Please try again")
