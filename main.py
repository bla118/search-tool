import search


try:
    # user input for path
    input_path = input("Enter the path to search from: ").strip()
    search_full_path = search.os.path.abspath(input_path)

    while not search.os.path.exists(search_full_path):
        input_path = input("Path does not exist, please enter a valid path: ")
        search_full_path = search.os.path.abspath(input_path)

    # user input for phrases to search for
    keywords = input("Enter the keywords to find (comma-separated): ").lower()
    print(f"Searching for {keywords}")
    keywords = [keyword.strip() for keyword in keywords.split(",")]
        
    # user input for the workbook's name and location
    input_workbook = input("Enter the name you would like for the Excel workbook (Default is FileSearch): ")
    if input_workbook.strip() == "":
        input_workbook = "FileSearch"

    workbook_xlsx = input_workbook + ".xlsx"
    input_dest_path = input("Enter the destination path you would like the file saved in: ").strip()

    dest_full_path = search.os.path.abspath(input_dest_path)
    # ensure that the user enters a valid paths
    while not search.os.path.exists(dest_full_path):
        input_dest_path = input("Path does not exist, please enter a valid path: ")
        dest_full_path = search.os.path.abspath(input_dest_path)

    while search.os.path.exists(dest_full_path + "/" + workbook_xlsx):
        input_workbook = input(f"{workbook_xlsx} already exists. Please choose another name: ")
        workbook_xlsx = input_workbook + ".xlsx"

    search.search_files(search_full_path, keywords, workbook_xlsx, dest_full_path)

except KeyboardInterrupt:
    search.key_quit()

except Exception as e:
    print(f"Error: {e}")