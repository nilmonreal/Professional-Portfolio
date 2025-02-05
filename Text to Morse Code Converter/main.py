from morse_code_dict import morse_code_dict


reverse_morse_code_dict = {value: key for key, value in morse_code_dict.items()}


choice = input("Choose an option: 1) Text to Morse  2) Morse to Text\n")

if choice == '1':
    
    user_input = input("Enter a string to convert to Morse code: ").upper()

    morse_output = []
    unsupported_characters = []

    for char in user_input:
        if char in morse_code_dict:
            morse_output.append(morse_code_dict[char])
        else:
            morse_output.append('?')  
            unsupported_characters.append(char)

   
    if unsupported_characters:
        print(f"Unsupported characters found: {', '.join(unsupported_characters)}")

    morse_string = ' '.join(morse_output)
    print("Morse Code:", morse_string)

elif choice == '2':
   
    morse_input = input("Enter Morse code to convert to text (use / for spaces between words): ").split(' ')

    decoded_output = []
    for code in morse_input:
        if code in reverse_morse_code_dict:
            decoded_output.append(reverse_morse_code_dict[code])
        elif code == '/':  
            decoded_output.append(' ')
        else:
            decoded_output.append('?')  

    decoded_string = ''.join(decoded_output)
    print("Decoded Text:", decoded_string)

else:
    print("Invalid choice. Please enter 1 or 2.")
