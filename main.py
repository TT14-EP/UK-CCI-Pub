import package
import package.pub_modules
import os

# Start-up interface
startinput = input('Insert 1 to start. Otherwise, insert 0 to exit:')
if startinput == '0':
    print("***** Program Exits. Goodbye Earthling! *****\n")
    sys.exit(0)
else:
    print("***** Program Starts *****\n")

    # Confirm/amend file names of Python datasheet and Word template (assume they are on the same directory as main.py)
    pyfnameinput = input('Input Python datasheet file name (including extension) and press Enter if you wish to amend. Leave blank and press Enter otherwise.:')
    if pyfnameinput == "":
        print("***** No amendment on Python datasheet file name. *****")
        print("***** Current Python datasheet file name: " + package.datasheetname + " *****\n")
    else:
        package.datasheetname = pyfnameinput
        package.datasheetpath = os.path.abspath(os.path.join(os.getcwd(), pyfnameinput))
        print("***** Python datasheet file name amended. *****")
        print("***** Current Python datasheet file name: " + package.datasheetname + " *****\n")

    wdtemplatefnameinput = input('Input Word template file name (including extension) and press Enter if you wish to amend. Leave blank and press Enter otherwise.:')
    if wdtemplatefnameinput == "":
        print("***** No amendment on Word template file name. *****")
        print("***** Current Word template file name: " + package.wdtemplatename + " *****\n")
    else:
        package.wdtemplatename = wdtemplatefnameinput
        package.wdtemplatepath = os.path.abspath(os.path.join(os.getcwd(), wdtemplatefnameinput))
        print("***** Python datasheet file name amended. *****")
        print("***** Current Python datasheet file name: " + package.wdtemplatename + " *****\n")


    # Get user input on classes to process
    final_list, df = package.pub_modules.row_picker(package.datasheetpath)
    completed = []

    # Core processes
    for row in final_list:
        print("********** Start processing " + str(row) + " **********\n")

        wordoutputname, wordoutputpath = package.pub_modules.generate_cci_summary(df, row, package.wdtemplatepath, package.pdffolderpath)

        package.pub_modules.insert_line_graph(df, row, package.plotfolderpath, wordoutputpath)

        package.pub_modules.word_to_pdf(package.pdffolderpath, wordoutputpath, wordoutputname)

        package.pub_modules.update_log(df, row, package.datasheetname)

        completed.append(row) #To record what has been completed.

        print("********** Finished processing " + str(row) + " **********\n")

    print('===================================================================================================================')
    print('Greetings Earthling. All processed: ', completed)
    print("\n***** Program Ends *****")
    print('===================================================================================================================')