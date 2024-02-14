from typing import List

class Account():
    def __init__(self, mail, key) -> None:
        self.mail = mail
        self.key = key

    def __str__(self) -> str:
        return f'{self.mail} {self.key}'

    def __repr__(self) -> str:
        return f'{self.mail} {self.key}'

def get_accounts(account_file) -> List[Account]:
    accounts = list()
    with open(account_file, "r", encoding="utf-8") as file:
        for line in file:
            parts = line.strip().split(":")

            if len(parts) >= 2:
                mail = parts[0].strip()
                key = parts[1].strip()
                accounts.append(Account(mail, key))
    return accounts


   
def register_stmp():
    pass

if __name__ == '__main__':


    accounts = get_accounts(account_file='accounts.txt')
    for account in accounts:
        print(account.mail, account.key)