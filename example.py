from Email_mod import return_service, gmail_send_message
import information

def main():

    service = return_service()
   
    # test = parse_excel_file('emails.xlsx',sheet_name="тест",column=1)
    
    test = ["armenovich2001@bk.ru"]
    gmail_send_message( service,
                        to_mail = test,
                        from_mail="jorj.knyazyan.15@gmail.com",
                        message_text=information.text, 
                        message_theme=information.theme,
                        )
                  
    
if __name__ == '__main__':
    main()
