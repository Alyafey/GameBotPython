import sys,os
from CandyCrush import CandyCrush 
from KeyHook import KeyHookMoniter

def main():   
    try:
        hook=KeyHookMoniter()
        candy = CandyCrush()
        hook.start(candy)
        candy.go()

    except Exception as e:
        print(e)
      

if __name__ == "__main__":

   main()