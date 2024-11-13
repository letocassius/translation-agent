import src.translation_agent.utils as ta

if __name__ == "__main__":
    test_text = "Hello, world!"
    try:
        translated_text = ta.translate("English", "Chinese", test_text, "China")
        print("Translated text:", translated_text)
    except Exception as e:
        print(f"Translation function error: {e}")