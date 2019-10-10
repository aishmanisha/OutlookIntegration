import requests


class OutlookService:
    def __init__(self):
        pass

    sender = ''
    receiver = ''
    subject = ''
    body = ''

    @staticmethod
    def read_emails(at):
        resp = requests.get('https://graph.microsoft.com/v1.0/me/messages?search="recipients:shekhar.pragati@outlook.com"', headers={"Authorization": at})
        if resp.status_code != 200:
            print ("error in api call")
        else:
            json_response = resp.json()
            email = json_response["value"][0]
            subject = email["subject"]
            body = email["bodyPreview"]
            sender = email["from"]["emailAddress"]["address"]

            print (subject)
            print (body)
            print(sender)

    @staticmethod
    def send_email(subject, reply , at):
        data = {
            "subject": subject,
            "importance": "normal",
            "body": {
                "contentType": "HTML",
                "content": reply
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": "shekharpragati143@gmail.com"
                    }
                }
            ]
        }
        print (data)
        draft_response = requests.post('https://graph.microsoft.com/v1.0/me/messages', json=data,
                                       headers={"Authorization": at})
        print (draft_response)
        if 200 <= draft_response.status_code < 300:
            json_response = draft_response.json()
            draft_id = json_response["id"]

            print ("draft_id => " + draft_id)

            email_response = requests.post("https://graph.microsoft.com/v1.0/me/messages/" + draft_id + "/send",
                                           headers={"Authorization": at})

            if 200 <= email_response.status_code < 300:
                print ("success in sending")
            else:
                print ("error in sending email")

            print (json_response)

        else:
            print ("error in api call")


if __name__ == '__main__':
    access_token = "Bearer EwBwA8l6BAAUO9chh8cJscQLmU+LSWpbnr0vmwwAAVnNvQWLzxHlH/La0xxMNYjExfMPLf0+B76pFdJ/psaIEBHo44hovUl2k62fk5EYi32pclyycYAxwv7T1Zqtzw+sTCEXiINRAl+j0TsQKmV4euptEe+D9pnSJfIiC2IWLV1B2+IULSQHgpqawpQQzAl6yzgCh6vurKqxWMU03lNoRK0yURokFiLbhTbz/ZUoAL6/uzmMhIyiCF2gf1mEHRVxHa+XIMupBlzGh1IJ2BExbffeK5geKu1PSjgRL3n0T67E+ErR7S1Ek2Gi1nRm3m9g9jcxpFRBQZolC7ZB5mbi8QSVDgRRkNXocl/ict00IEUjLh3p4vrVcfl55scsiiQDZgAACKt5RfJ06FL/QAKjGvb8sDCx1GUqNP3PvXw2nod2IUZ7znTQIJ76vyCXge5hmbeF1ntlnXo1RkGCQUEian5+HHcEt0OiQDVK5O5ov2em5ouEoSATXs2WiH7678FFONPwQHBuKw8EwlShHWHKX0dfehMBJ7eOC2XjR9Bu1UBc9nhMuhqmmX3/+bcicoKyozF7WFtPyCDwTccsS1hdtymHvSI7KUo8aixO7XXH6SaQl+34XaAJIanNReQlmBkuyBnperY7WEElF+lL/KqErHidGSboilJ1mR7oRO7bSAaMFvYnPeR8bUmkgrgpoRHZj7EaCdjIMlnHIEeRYd2cWjdXE28LaG3MR4Kk/etwh6vhvR5KVtTja4cCVlFoURtqcjo4muZUcJi/p15mBx7We2rSDb59jVvXx6ND/y2mrd1PSxcfNKGqO/aCB3IUIVYwHTuZS+g487pTZ42FbkUkjfA1CK945B6k/e/M0LmkSH/hN5anoA27QhZP+QcLhEGBnLpzdGviiR7jRy05YFYpzHFrYnxnQD3z8X2J9W4NFb5bUA5YNbe87DNP2tHvNjQq4IVVDuxcggYfSeiYwLzMB7tlF2a0v+/nKZcAbLF3YOWlhSYelEBqdeJjc2TCUyeYXCqvLTa+/4aZANFFmPZU9q2Khku4FRhkQMQnH2VfNpJpcfLqYEVNNU11v0+uDSNh39gIMKxkR9JMYqn2w1sytdV6T1v11J9ixy30HGSYkYJnqvw7Z/5dVbDuNsmGEb/9hNnd1nyO642m79jChGyGAg=="
    OutlookService.send_email("Subject of received email", "Reply text from computing engine" , access_token)
    # OutlookService.read_emails(access_token)
