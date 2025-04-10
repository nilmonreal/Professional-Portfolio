class TicTacToe:
    def __init__(self):
        # Initialize the board with 9 empty spaces
        self.board = [' ' for _ in range(9)]
        # Current player starts as 'X'
        self.current_player = 'X'

    def print_board(self):
        """Print the current state of the board."""
        for i in range(0, 9, 3):
            print(f' {self.board[i]} | {self.board[i+1]} | {self.board[i+2]} ')
            if i < 6:
                print('-----------')

    def is_valid_move(self, position):
        """Check if the selected position is valid."""
        return 0 <= position < 9 and self.board[position] == ' '

    def make_move(self, position):
        """Place the current player's mark on the board."""
        if self.is_valid_move(position):
            self.board[position] = self.current_player
            return True
        return False

    def check_winner(self):
        """Check if there's a winner."""
        # Winning combinations
        win_combinations = [
            # Rows
            (0, 1, 2), (3, 4, 5), (6, 7, 8),
            # Columns
            (0, 3, 6), (1, 4, 7), (2, 5, 8),
            # Diagonals
            (0, 4, 8), (2, 4, 6)
        ]

        for combo in win_combinations:
            if (self.board[combo[0]] == self.board[combo[1]] == self.board[combo[2]] != ' '):
                return self.board[combo[0]]
        
        return None

    def is_board_full(self):
        """Check if the board is full (a draw)."""
        return ' ' not in self.board

    def switch_player(self):
        """Switch the current player."""
        self.current_player = 'O' if self.current_player == 'X' else 'X'

    def play(self):
        """Main game loop."""
        print("Welcome to Tic Tac Toe!")
        print("Players take turns marking spaces. X goes first.")
        print("Enter a number 0-8 corresponding to the board positions:")
        print(" 0 | 1 | 2 ")
        print("-----------")
        print(" 3 | 4 | 5 ")
        print("-----------")
        print(" 6 | 7 | 8 ")

        while True:
            # Print current board state
            self.print_board()
            
            # Prompt for move
            print(f"Player {self.current_player}'s turn")
            try:
                move = int(input("Enter your move (0-8): "))
            except ValueError:
                print("Invalid input. Please enter a number between 0 and 8.")
                continue

            # Check and make move
            if self.make_move(move):
                # Check for winner
                winner = self.check_winner()
                if winner:
                    self.print_board()
                    print(f"Player {winner} wins!")
                    break
                
                # Check for draw
                if self.is_board_full():
                    self.print_board()
                    print("It's a draw!")
                    break
                
                # Switch players
                self.switch_player()
            else:
                print("Invalid move. Try again.")

# Run the game
if __name__ == "__main__":
    game = TicTacToe()
    game.play()