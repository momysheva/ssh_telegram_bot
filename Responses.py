def sample_responces(input_text):
    user_msg=str(input_text).lower()

    if user_msg in ("hello", "hi"):
        return "Please send all needed documents!"
    
    return "I don't understand"
