import random
import BJExceptions

#Card class that gives you a String Representaiton of it's self
#And can be init to differint cards of 1 - Ace and D, H, S, C.
class Playing_Card:
    """Card Class that can hold a card value and a suit"""

    def __init__(self, value, suit):
        self.value = value
        self.suit = suit

    def card_BJvalue(self):
        """Return's the Card's Value in BlackJack."""
        try:
            return int(self.value)
        except:
            if self.value == 'K' or self.value == 'Q' or self.value == 'J':
                return 10
            elif self.value == 'A':
                return 11

    def is_Ace(self):
        """Returns True is the card is an Ace. False if not."""
        if self.value == 'A':
            return True
        else:
            return False

    def __str__(self):
        return f'{self.value}{self.suit}'
#Deck class that holds 52 or fewer cards.
#when it is init it makes 52 cards and is shuffled.
#can shuffle it's self and pop you the card on top.
class Deck_Of_Cards:
    """
    A deck of 52 cars init to have all the cards.
    can be shuffled and can pop the card on top.
    """

    def __init__(self):
        self.deck = []
        for suit in [u'\u2663', u'\u2660', u'\u2665', u'\u2666']:
            for card_value in [2,3,4,5,6,7,8,9,10,'J','Q','K','A']:
                self.deck += [Playing_Card(card_value, suit)]

        random.shuffle(self.deck)

    def shuffle_deck(self):
        """shuffles the deck of cards."""
        random.shuffle(self.deck)

    def deck_length(self):
        """Returns the Number of remaining cards."""
        return len(self.deck)

    def get_top_card(self):
        """Removes the Last card from the deck and returns it."""
        return self.deck.pop()

    def __str__(self):
        return f'A deck of card\'s with {len(self.deck)} cards in it.'


#Person Class. Base for Dealer and Player
class Person:
    def __init__(self):
        self.hand = []
        self.is_bust = False

    def add_card_to_hand(self, card):
        """adds a card to the person\'s hand"""
        if type(card) == Playing_Card:
            self.hand += [card]
        else:
            raise Exception('Tried to add something other than a Playing Card to the player\'s hand')

    def count_hand_value(self):
        """Counts the value of the Players hand."""
        sum = 0
        ace = 0
        for card in self.hand:
            if card.is_Ace():
                ace += 1
            sum += card.card_BJvalue()
        if sum > 21 and ace > 0:
            sum -= 10
        if sum > 21 and ace > 1:
            sum -= 10
        if sum > 21 and ace > 2:
            sum -= 10
        if sum > 21 and ace > 3:
            sum -= 10
        return sum

    def has_blackjack(self):
        """Test to see if the person\'s hand is a blackjack. Returns a boolean."""
        if self.count_hand_value() == 21 and len(self.hand) == 2:
            return True
        else:
            return False

#Dealer Class
class Dealer(Person):
    """A Dealer Class to hold the dealer\'s hand."""
    def __init__(self):
        Person.__init__(self)
        self.hidden = True


    def __str__(self):
        if self.hidden:
            hidden_hand = ['XX']
            for card in self.hand[1:]:
                hidden_hand += [str(card)]
            return str(hidden_hand)
        else:
            printable_hand = []
            for card in self.hand:
                printable_hand += [str(card)]
            return str(printable_hand)
#Player Class has name and Money balance
class Player(Person):
    """A Player with a name and a money balance and a hand of cards."""

    def __init__(self, player_name, balance):
        Person.__init__(self)
        self.name = player_name
        self.balance = balance

    def empty_hand(self):
        self.hand = []

    def withdraw(self, amt):
        """take's money from the Player's balance if sufficient funds."""
        if amt >= 0 and self.balance >= amt:
            self.balance -= amt
        elif amt <= 0:
            raise BJExceptions.CheatingBetError()
        else:
            raise BJExceptions.InsufficientFundsError()

    def deposit(self, amt):
        """Put's money into the players balance"""
        self.balance += amt

    def __str__(self):
        printable_hand = []
        for card in self.hand:
            printable_hand += [str(card)]
        return str(printable_hand)
