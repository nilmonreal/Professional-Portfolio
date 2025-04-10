import turtle
import random

class Breakout:
    def __init__(self):
        # Screen setup
        self.screen = turtle.Screen()
        self.screen.title("Breakout")
        self.screen.bgcolor("black")
        self.screen.setup(width=800, height=600)
        self.screen.tracer(0)  # Turn off automatic screen updates

        # Paddle
        self.paddle = turtle.Turtle()
        self.paddle.shape("square")
        self.paddle.color("white")
        self.paddle.shapesize(stretch_wid=1, stretch_len=5)
        self.paddle.penup()
        self.paddle.goto(0, -250)

        # Ball
        self.ball = turtle.Turtle()
        self.ball.shape("circle")
        self.ball.color("white")
        self.ball.penup()
        self.ball.goto(0, -230)
        self.ball_dx = 0.3
        self.ball_dy = 0.3

        # Bricks
        self.bricks = []
        self.create_bricks()

        # Score
        self.score = 0
        self.score_display = turtle.Turtle()
        self.score_display.color("white")
        self.score_display.penup()
        self.score_display.hideturtle()
        self.score_display.goto(0, 260)
        self.update_score()

        # Lives
        self.lives = 3
        self.lives_display = turtle.Turtle()
        self.lives_display.color("white")
        self.lives_display.penup()
        self.lives_display.hideturtle()
        self.lives_display.goto(-380, 260)
        self.update_lives()

        # Game over display
        self.game_over_display = turtle.Turtle()
        self.game_over_display.color("white")
        self.game_over_display.penup()
        self.game_over_display.hideturtle()

        # Keyboard bindings
        self.screen.listen()
        self.screen.onkeypress(self.move_paddle_left, "Left")
        self.screen.onkeypress(self.move_paddle_right, "Right")

        # Game state
        self.game_active = True

    def create_bricks(self):
        """Create a grid of bricks"""
        colors = ["red", "orange", "green", "yellow"]
        for y in range(4):
            for x in range(11):
                brick = turtle.Turtle()
                brick.shape("square")
                brick.color(colors[y])
                brick.shapesize(stretch_wid=1, stretch_len=4)
                brick.penup()
                brick.goto(-380 + x * 76, 230 - y * 30)
                self.bricks.append(brick)

    def move_paddle_left(self):
        """Move paddle left"""
        x = self.paddle.xcor()
        if x > -350:
            self.paddle.setx(x - 20)

    def move_paddle_right(self):
        """Move paddle right"""
        x = self.paddle.xcor()
        if x < 350:
            self.paddle.setx(x + 20)

    def update_score(self):
        """Update and display the score"""
        self.score_display.clear()
        self.score_display.write(f"Score: {self.score}", 
                                 align="center", 
                                 font=("Courier", 16, "normal"))

    def update_lives(self):
        """Update and display lives"""
        self.lives_display.clear()
        self.lives_display.write(f"Lives: {self.lives}", 
                                 align="left", 
                                 font=("Courier", 16, "normal"))

    def reset_ball(self):
        """Reset ball position and give it a random initial direction"""
        self.ball.goto(0, -230)
        self.ball_dx = random.choice([-0.3, 0.3])
        self.ball_dy = 0.3

    def game_over(self, result):
        """Handle game over conditions"""
        self.game_active = False
        self.ball.hideturtle()
        self.paddle.hideturtle()
        
        # Clear all bricks
        for brick in self.bricks:
            brick.hideturtle()
        
        # Display game over message
        self.game_over_display.goto(0, 0)
        self.game_over_display.write(result, 
                                     align="center", 
                                     font=("Courier", 24, "normal"))

    def play(self):
        """Main game loop"""
        while self.game_active:
            # Update screen
            self.screen.update()

            # Move the ball
            self.ball.setx(self.ball.xcor() + self.ball_dx)
            self.ball.sety(self.ball.ycor() + self.ball_dy)

            # Border checking
            # Left and right walls
            if self.ball.xcor() > 390 or self.ball.xcor() < -390:
                self.ball_dx *= -1

            # Top wall
            if self.ball.ycor() > 290:
                self.ball_dy *= -1

            # Bottom wall (lose a life)
            if self.ball.ycor() < -290:
                self.lives -= 1
                self.update_lives()
                
                if self.lives > 0:
                    self.reset_ball()
                else:
                    self.game_over("Game Over!")
                    break

            # Paddle collision
            if (self.ball.ycor() < -240 and 
                self.paddle.xcor() - 50 < self.ball.xcor() < self.paddle.xcor() + 50):
                self.ball_dy *= -1
                # Add some angle variation based on where the ball hits the paddle
                angle_modifier = (self.ball.xcor() - self.paddle.xcor()) / 50
                self.ball_dx += angle_modifier * 0.2

            # Brick collision
            for brick in self.bricks[:]:
                if (abs(self.ball.xcor() - brick.xcor()) < 40 and 
                    abs(self.ball.ycor() - brick.ycor()) < 20):
                    # Reverse ball direction
                    self.ball_dy *= -1
                    
                    # Remove brick
                    brick.hideturtle()
                    self.bricks.remove(brick)
                    
                    # Increase score
                    self.score += 10
                    self.update_score()

            # Check if all bricks are destroyed
            if not self.bricks:
                self.game_over("You Win!")
                break

        # Keep window open
        self.screen.mainloop()

def main():
    game = Breakout()
    game.play()

if __name__ == "__main__":
    main()