#Set up global variables
playing = True #used to exit while loop.
global turncount
p1turn = True #used to tell whoes turn it is. False means it's player 2's turn.
b = [1,2,3,4,5,6,7,8,9] #initalize the answer board
#define some commonly used fucntions.
#Prints the current board.

def print_board():
    #print("\u0332".join('{0}|{1}|{2}\n{3}|{4}|{5} '.format(b[0],b[1],b[2],b[3],b[4],b[5])) + '\n {0} | {1} | {2}  '.format(b[6],b[7],b[8]))
    print('_{0}_|_{1}_|_{2}_\n_{3}_|_{4}_|_{5}_'.format(b[0],b[1],b[2],b[3],b[4],b[5]) + '\n {0} | {1} | {2}  '.format(b[6],b[7],b[8]))
    pass

#check that a players move is valid and return a boolean
def validate_move(move):
    if move.isdigit():
        move = int(move)
        if 0 < move < 10:
            if b[move-1] == 'X' or b[move-1] == 'O':
                return False
            else:
                return True
        else:
            return False
    else:
        return False

#Check's the board for winning Conditions. returns the following.
#0 = continue playing
#1 = player 1 Wins
#2 = player 2 Wins
#3 = CAT
def check_board():
    global b
    global turncount

    if b[0] == b[1] == b[2] or \
    b[3] == b[4] == b[5] or \
    b[6] == b[7] == b[8] or \
    b[0] == b[3] == b[6] or \
    b[1] == b[4] == b[7] or \
    b[2] == b[5] == b[8] or \
    b[0] == b[4] == b[8] or \
    b[2] == b[4] == b[6]:
        if p1turn: #Then Player 2 made the winning move
            return 2
        else: #Player 1 made the winning move
            return 1
    elif turncount >= 9 : #max number of moves taken.
        return 3
    else:#continue playing
        return 0

#Welcome text to begin game and ask for names
print('Welcome!')
print('Let\'s play Tic Tac Toe!')
print('You know the rules! first player to get three marks in a row wins! \nAnd if neither player succeeds it is a CAT!')
player1 = input('Player 1, what is your name?')
print(f'Hello {player1}!')
player2 = input('Player 2 what is your name?')
print(f'Hello {player2}!\n')



#begin the game
turncount = 1
while playing:
    print_board()
    if p1turn: #player1's turn
        move = input(f'{player1}, where would you like to play?')
        if validate_move(move):
            move = int(move)
            b[move-1] = 'X'
            p1turn = False
        else:
            print('Sorry, that is an invalid move.\nTry entering one of the numbers of an available square.')
            continue
    else: #player 2's turn
        move = input(f'{player2}, where would you like to play?')
        if validate_move(move):
            move = int(move)
            b[move-1] = 'O'
            p1turn = True
        else:
            print('Sorry, that is an invalid move.\nTry entering one of the numbers of an available square.')
            continue

    b_cond = check_board()
    if b_cond == 0:
        turncount += 1
        continue
    elif b_cond == 1:
        print_board()
        print(f"Congratualtions {player1}, you won!")
        playing = False
        break
    elif b_cond == 2:
        print_board()
        print(f"Congratualtions {player2}, you won!")
        playing = False
        break
    elif b_cond == 3:
        print_board()
        print(f"CAT! That means you both loose. \n(But {player1} looses a bit more than {player2} because they had the advantage.)")
        playing = False
        break
    else:
        print('Error checking the board. Game Over.')
        playing = False
        break

#End while loop Code Block
