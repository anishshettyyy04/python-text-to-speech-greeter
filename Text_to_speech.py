import win32com.client as wincl

speaker_number = 0
speaker = wincl.Dispatch("SAPI.SpVoice")
voices = speaker.GetVoices()

speaker.Speak(f"Hello, my name is {voices.Item(speaker_number).GetAttribute('Name')}")
print("Type the , the of greetings to the person you want to give like Good Morning ,good afternoon , good Evening, good night \n ")
wish=input("Enter the Greeting : ")
speaker.Speak("Type the , the of greetings to the person you want to give like Good Morning ,good afternoon , good Evening, good night \n ")
wish=input("Enter the Greeting : ")
name_list = []
n = int(input("Enter the number of people to greet: "))
speaker.Speak("Enter the number of people to greet: ")

for i in range(n):
    speaker.Speak("Enter the name: ")
    names = input("Enter the name: ")
    name_list.append(names)

for name in name_list:
    speaker.Speak(f"{wish} , {name}!")
