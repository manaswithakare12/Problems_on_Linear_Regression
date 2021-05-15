""" 
          ^    ^    ^    ^    ^    ^    ^    ^    ^    ^    ^    ^    ^    ^
         /B\  /e\  /t\  /t\  /e\  /r\  /T\  /e\  /a\  /c\  /h\  /i\  /n\  /g\
        <___><___><___><___><___><___><___><___><___><___><___><___><___><___>

Developed by: Ashish Kumar
"""

from src import *
import os

if __name__ == "__main__":

    s3Upload = int(
        input(
            "\nDo you want to upload the video to s3?[Press 1 for Yes and 0 if you have Video url, (default 0)]:"
        )
        or "0"
    )

    if s3Upload == 1:
        try:
            upload_folder_to_s3(s3.Bucket(Bucket_name), path + "/videos/", Folder_in_S3)
        except:
            print(f"\nUpload fail")
    else:
        print("Proceeding to next step...")

    file_names = os.listdir(path + "/videos/")
    job_uri = str(
        input(
            f"\nURL of vidio uploaded in s3 [default: https://{Bucket_name}.s3.{region_name}.amazonaws.com/{Folder_in_S3}/{file_names[0]} ]:"
        )
        or f"https://{Bucket_name}.s3.{region_name}.amazonaws.com/{Folder_in_S3}/{file_names[0]}"
    )

    try:
        download(job_uri, LanguageCode="en-US")
        print("Check `./out/folder for transcription`")
    except:
        print("Error Occured !!")
