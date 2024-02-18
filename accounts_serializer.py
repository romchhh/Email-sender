from typing import List

class Account():
    def __init__(self, mail, key, banned = False) -> None:
        self.mail = mail
        self.key = key
        self.banned = banned

    def __str__(self) -> str:
        return f'{self.mail} {self.key} banned:{self.banned}'

    def __repr__(self) -> str:
        return f'{self.mail} {self.key} banned:{self.banned}'

def get_accounts(account_file) -> List[Account]:
    accounts = list()
    with open(account_file, "r", encoding="utf-8") as file:
        for line in file:
            parts = line.strip().split(":")

            if len(parts) >= 2:
                mail = parts[0].strip()
                key = parts[1].strip()
                banned = False
                if len(parts) >= 3 and parts[2].strip(): 
                    banned = True
                accounts.append(Account(mail, key, banned))
    return accounts
    

def became_banned():
    pass


if __name__ == '__main__':


    accounts = get_accounts(account_file='accounts.txt')
    for account in accounts:
        print(account)
