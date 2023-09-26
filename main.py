import importlib
from win32com.client import Dispatch
import random

while True:
    Person1=input("Enter the name of person1 ")
    Person2=input("Enter the name of person2 ")
    Person3=input("Enter the name of person3 ")
    animal=input("Enter the name of any animal ")
    place=input("Enter the name of any place ")

    choice=int(input("""
Select your options:
1---------->Read the story.
2---------->Listen to the story.   """))
    select_story=random.randint(1,4)
    story=f"story{select_story}"
    module = importlib.import_module(story)
    match choice:
        case 1:
            story=module.makeStory(Person1,Person2,Person3,animal,place)
            print("Enjoy the story")
            print(story)
    
        case 2:
            sepak=Dispatch("SAPI.SpVoice")
            sepak.Speak("Listen the story carefully.")
            sepak.Speak(module.makeStory(Person1,Person2,Person3,animal,place))
    
    response=int(input('''Select one of the following:
1-------------->Play Again.
0-------------->Exit.    '''))
    if not response:
        break   