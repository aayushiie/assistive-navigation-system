import win32com.client

def t2s():       
    # initialise
    speaker = win32com.client.Dispatch("SAPI.SpVoice")

    # Configure voice settings
    speaker.Rate = 0      # Speech speed (-10 to 10, 0 is normal)
    speaker.Volume = 100  # Volume (0 to 100)

    while True:
        user_input = input("Enter text: ").strip()

        if user_input.lower() in ['quit', 'exit', 'q']:
            print("\nGoodbye!")
            break
        
        if not user_input:
            continue
        
        print(f"Speaking: {user_input}")
        speaker.Speak(user_input)

    speaker = None

t2s()