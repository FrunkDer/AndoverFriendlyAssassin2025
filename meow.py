import win32com.client
import csv
import random

outlook = win32com.client.Dispatch("Outlook.Application")

# mail = outlook.CreateItem(0)
# mail.To = "jinfanfrank@gmail.com"
# mail.Subject = "Automated Email"
# mail.Body = "This is an automated email sent from Outlook using Python."
# mail.Send()
# print("Email sent successfully!")

def makepairs():
    with open(r'C:\Users\jinfa\OneDrive\Desktop\Friendly Assassin\Friendly Assassin Signups.csv', newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)  # Create a reader object
        nested_list = [row for row in reader]  # Convert reader object to list
        for littlelist in nested_list:
            littlelist.pop(0)
            littlelist.pop()

        random.shuffle(nested_list)
        pairs = list(zip(nested_list[::2], nested_list[1::2]))
        
        return pairs

def write_pairs_to_csv(pairs, filename="pairs.csv"):
    with open(filename, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        for pair1, pair2 in pairs:
            full_name1 = f"{pair1[0]} {pair1[1]}"
            full_name2 = f"{pair2[0]} {pair2[1]}"
            writer.writerow([f"{full_name1},{full_name2},{pair1[2]},{pair2[2]}"])  # Write names in required format    

def emailpair(task, filename="pairs.csv"):
    with open(filename, mode='r', newline='', encoding='utf-8') as file:
        reader = csv.reader(file)
        for row in reader:
            rowlist = row[0].split(",")
            sendmessage(rowlist, task)

def sendmessage(row, task):
    message = f'''Hi {row[0]},
    Your partner will be {row[1]}. 
    
    Your task is as follows: 
    {task}
    
    Remember to DM us the picture at @andover.2026 with your names to get credit for the task! The earlier you complete it, the more points you get. 
    Good luck!
    Frank'''
    mail = outlook.CreateItem(0)
    mail.To = row[2]
    mail.Subject = "FRIENDLY ASSASSIN ASSIGNMENT"
    mail.Body = message
    mail.Send()

    message = f'''Hi {row[1]},
    Your partner will be {row[0]}. 
    
    Your task is as follows: 
    {task}
    
    Remember to DM us the picture at @andover.2026 with your names to get credit for the task! The earlier you complete it, the more points you get. 
    Good luck!
    Frank'''
    mail = outlook.CreateItem(0)
    mail.To = row[3]
    mail.Subject = "FRIENDLY ASSASSIN ASSIGNMENT"
    mail.Body = message
    mail.Send()
    
def STARTTASK(task):
    pairs = makepairs()
    write_pairs_to_csv(pairs)
    emailpair(task)
    print("Successfully Started Task!")

import csv

def addscore(first, last, place):
    filename = "Friendly Assassin Signups.csv"
    updated_rows = []
    found = False

    # Read the CSV file and modify the score if the name matches
    with open(filename, mode='r', newline='', encoding='utf-8') as file:
        reader = csv.reader(file)
        for row in reader:
            if row[1] == first and row[2] == last:
                score = int(row[4])  # Convert score to integer
                row[4] = str(score + placecalc(place))  # Update score
                print(f"Added Score to {first} {last}")
                found = True
            updated_rows.append(row)

    if not found:
        print("Name not found")
        return False

    # Write the modified data back to the CSV file
    with open(filename, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerows(updated_rows)  # Write all rows back

    return True


def placecalc(place):
    if place == 1:
        return 10
    if place < 6:
        return 5
    if place < 11:
        return 3
    if place > 10:
        return 1

def emailscores(note):
    emailstr = ''
    scorelist = []
    with open("Friendly Assassin Signups.csv", mode='r', encoding='utf-8') as f:
        reader = csv.reader(f)
        for row in reader:
            emailstr += f"{row[3]}; "
            if (int)(row[4]) != 0:
                scorelist.append([row[1], row[2], (int)(row[4])])
        
    scorelist.sort(key=lambda x: x[2], reverse=True)

    message = f'''Hey Friendly Assassins,
    {note}
    Keep your eyes peeled for a new task soon. For now, here's the leaderboard:
    
{leaderboardStr(scorelist)}
Frank
    '''
    mail = outlook.CreateItem(0)
    mail.To = emailstr
    mail.Subject = "LEADERBOARD UPDATE"
    mail.Body = message
    mail.Send()
    print("Email Sent Successfully")


def leaderboardStr(scores):
    s = ""
    for i, score in enumerate(scores):
        s += f"{i+1}. {score[0]} {score[1]}: {score[2]} Point{'s' if score[2] != 1 else ''}!\n"
    return s



