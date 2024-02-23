import string

card = "card1,card2,card3"
cards = card.split(",")
for i in range(1, 4):
    print(cards[i - 1])