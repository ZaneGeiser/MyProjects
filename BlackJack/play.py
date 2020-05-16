import BJClasses
import BJExceptions
import time

def print_playing_field(player, dealer):
    """takes the player\'s hand and dealer\'s hand and prints the playing field."""
    print()
    print(f'Dealer:\n   {dealer}')
    print(f'{player1.name}:\n   {player1}')
    print()
    pass

def clear_screen():
    """Prints a bunch of lines to clear the terminal screen."""
    print('\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n')
    pass

def take_bet(person):
    betting = True
    while betting:
        bet = input(f'Hey {person.name}, how much would you like bet this round?')
        try: #to process the players bet and ask again if they gave an invalid response.
            bet = int(bet)
            person.withdraw(bet)
            betting = False
            return bet
            break
        except BJExceptions.InsufficientFundsError:
            print(f"Sorry. You only have {person.balance} to bet so you can't afford that.")
            continue
        except BJExceptions.CheatingBetError:
            print("Betting <= 0 is cheating. Try again.")
            continue
        except:
            print(f"Sorry, \"{bet}\" is an invalid bet. Try using a number.")
            continue

def hitting_loop(person, dealer, deck):
    hitting = True
    while hitting and not person.is_bust:
        try:
            response = input("Would you like to hit or stay?")
            if response.lower().startswith('h'):
                #hit stuff
                person.add_card_to_hand(deck.get_top_card())
                if person.count_hand_value() > 21:
                    person.is_bust = True
                    hitting = False
            elif response.lower().startswith('s'):
                hitting = False
                #pass stuff
                pass
            else:
                print('invalid response. Try responding with \'h\' or \'s\'.')
                time.speep(1)
                continue
        except:
            print('there was an error with that')
            time.sleep(5)
        finally:
            print('\n\n')
            print_playing_field(person, dealer)
            time.sleep(1)
def dealer_turn(person, dealer, deck):
    hitting = True
    while hitting and not dealer.is_bust:
        if dealer.count_hand_value() < 18:
            print('\n\n')
            print("The dealer hits.")
            time.sleep(2)
            dealer.add_card_to_hand(deck.get_top_card())
            print_playing_field(person, dealer)
            time.sleep(3)
        else:
            print('\n\n')
            print("The dealer stays.")
            time.sleep(2)
            hitting = False
            break
        if dealer.count_hand_value() > 21:
            dealer.is_bust = True
            hitting = False
            break

def play_round(player1):
    player1 = player1
    player1.is_bust = False

    bet = take_bet(player1)
    #Inilize the playing field and deal cards.
    deck = BJClasses.Deck_Of_Cards()
    dealer = BJClasses.Dealer()
    player1.add_card_to_hand(deck.get_top_card())
    dealer.add_card_to_hand(deck.get_top_card())
    print_playing_field(player1, dealer)
    time.sleep(1)
    player1.add_card_to_hand(deck.get_top_card())
    print_playing_field(player1, dealer)
    time.sleep(1)
    dealer.add_card_to_hand(deck.get_top_card())
    print_playing_field(player1, dealer)
    time.sleep(1)

    #begin the player's turn
    hitting_loop(player1, dealer, deck)

    #begin the dealer's turn
    if not player1.is_bust:
        dealer_turn(player1, dealer, deck)

    #Display the winner
    clear_screen()
    dealer.hidden = False
    print("It's time to reveal the hand!")
    time.sleep(1)
    print_playing_field(player1, dealer)
    time.sleep(1)
    if player1.is_bust:
        print(f'{player1.name} has bust! the dealer wins this round and you lose your bet of {bet}')
        print(f'Your new balance is: {player1.balance}.')
    elif dealer.is_bust:
        print(f'The dealer has bust! You won ${bet * 2}')
        player1.deposit(bet * 2)
        print(f'Your new balance is: {player1.balance}.')
    elif player1.has_blackjack() and dealer.has_blackjack():
        print('You and the dealer have tied! this round is a push and your bet is returned to you.')
        player1.deposit(bet)
        print(f'Your balance is: {player1.balance}.')
    elif player1.has_blackjack():
        print('WINNER WINNER CHICKEN DINNER! \nYou have won this round with blackjack!')
        player1.deposit(bet * 2)
        print(f'Your new balance is: {player1.balance}.')
    elif dealer.has_blackjack():
        print(f'The Dealer has won with a BLACKJACK! You lose your bet of {bet}')
        print(f'Your new balance is: {player1.balance}.')
    elif player1.count_hand_value() == dealer.count_hand_value():
        print('You and the dealer have tied! this round is a push and your bet is returned to you.')
        player1.deposit(bet)
        print(f'Your balance is: {player1.balance}.')
    elif player1.count_hand_value() > dealer.count_hand_value():
        print(f'You have won with {player1.count_hand_value()} points! You won ${bet * 2}')
        player1.deposit(bet * 2)
        print(f'Your new balance is: {player1.balance}.')
    else:
        print(f'The Dealer has won with {dealer.count_hand_value()} points! You lose your bet of {bet}')
        print(f'Your new balance is: {player1.balance}.')

if __name__ == '__main__':
    print("This is a BlackJack Game. If you don't know how to play, google it.")
    print("WELCOME!")
    print("Only one player has been detected.")
    name = input("One Payer, what would you like to introduce yourself as?")
    print(f'Okay, {name}. I can call you that if you like.')
    balance = input("How much money are you bringing to this game?")

    try:
        balance = int(balance)
        if balance <= 0:
             raise Exception()
        player1 = BJClasses.Player(name, balance)
    except:
        print("Nah. You can start with $100.")
        player1 = BJClasses.Player(name, 100)

    round_count = 1
    playing = True
    while playing:
        time.sleep(2)
        clear_screen()
        player1.empty_hand()
        print(f'Begin round {round_count}')
        play_round(player1)
        if player1.balance == 0:
            print('Looks like the house has cleaned you out. \nThe casino has no more use for you and are escourting you out. \nHave a nice day!')
            playing = False
            continue
        responding = True
        while responding:
            response = input('Would you like to play another round?')
            if response.lower().startswith('y'):
                round_count += 1
                break
            elif response.lower().startswith('n'):
                print("Okay, have a great day and don't spend it all in one place!")
                playing = False
                break
            else:
                print(f'Sorry. \"{response}\" is an invalid response. Try saying yes or no.')
                continue
