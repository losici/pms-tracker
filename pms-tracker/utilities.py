def input_in_range(prompt, min_value, max_value):
    while True:
        try:
            user_input = input(prompt)
            if not user_input:
                user_input = 0
            else:
                user_input = int(user_input)
            if min_value <= user_input <= max_value:
                return user_input
            print(f"Input must be between {min_value} and {max_value}. Try again.")
        except ValueError:
            print("Invalid input! Please enter a valid integer.")

if __name__=="__main__":
    result = input_in_range("Do a self-test, insert an input between 0 and 5:\n", 0, 5)
    print(result)