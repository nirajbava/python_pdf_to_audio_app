import comtypes.client
import PyPDF2
import random



while(True):


    def pdf_to_audio():
        # read pdf 
        try:
            pdf_file_name = raw_input("Enter Your File Name (ex: book) : ")
        except:
            print("try again")

        book=""

        try: 
            book = open(pdf_file_name, 'rb')
        except Exception as a:
            pdf_file_name = raw_input("Enter Your File Name (ex: book) : ")
            print("try aging")

        pdfReader = PyPDF2.PdfFileReader(book)
        pages = pdfReader.numPages

        print('''
        **********WELCOME TO OUR APP**********''')

        # get user input
        page_to_start = int(input('''
        Enter Page Number To play From : '''))

        page_to_end  = int(input('''
        Enter page Number To End : '''))


        # make file name
        rand = random.randint(1,1000)

        file_name= "D:\\audio_to_pdf_app\\audio_book_audio_files\\audio" + str(rand) + ".mp3"

        my_text =""

        # set page number to read 
        for num in range(page_to_start-1, page_to_end):
            page = pdfReader.getPage(num)
            my_text = page.extractText()

            print('''
            %%%%%%% CREATING YOUR AUDIO FILE %%%%%%%
            ''')

            # add comtypes module
            speak = comtypes.client.CreateObject("SAPI.SpVoice")
            filestream = comtypes.client.CreateObject("SAPI.spFileStream")

            # create audio file
            filestream.open(file_name, 3, False)
            speak.AudioOutputStream = filestream 
            speak.Rate = -2
            speak.Volume = 100
            speak.Speak(my_text)
            filestream.close()

            print('''
            ######FILE CREATED SUCCESSFULLY#######
            ''')

    pdf_to_audio()