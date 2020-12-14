#TODO: Create a letter using starting_letter.docx 

import docx

def extract_names():
    names_x = []
    with open("./Input/Names/invited_names.txt", mode="r") as name_file:
        names_x = name_file.readlines()
    for i_x in range(len(names_x)):
        name_x = names_x[i_x]
        name_x = name_x.replace("\n", "").strip()
        names_x[i_x] = name_x
    return names_x

def mail_merge():
    names = extract_names()

    name_marker_text = "<name>"

    input_document = docx.Document("./Input/Letters/starting_letter.docx")



    strFilename_Output_Base = "./Output/ReadyToSend/"

    for i in range(len(names)):
        strName = names[i].title()
        strFilename_Output = strFilename_Output_Base + "Letter_to_" + strName + ".docx"

        # Replace name
        output_document = input_document

        for paragraph in output_document.paragraphs:
            if name_marker_text in paragraph.text:
                strParagraph = paragraph.text.replace(name_marker_text, strName)
                paragraph.text = strParagraph

        # Output Letter
        output_document.save(strFilename_Output)


# Run Mail Merge Program
mail_merge()



