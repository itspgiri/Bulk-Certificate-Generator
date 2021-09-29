# Bulk-Certificate-Generator
Let's anyone create as many certificates in a fraction of time with just small input to specify where to print the names, data, etc. I have writen a step-by-step process to help run this code with minimal changes in the code - as the certificate will change person to person and so will the code. It's a simple code, for anyone even without a python background to run and implement.

# INTRODUCTION
It's a basic python program based on the Pillow Module to create bulk certificates for any event. Most softwares available online render a low quality certificate. It only requires you to make two changes

- Update the excel sheet with your names for the Certificate.
- Update where in the certificate should the name be printed.

# PROCESS
Making it as detailed as possible so even a beginner can complete this process. Hence, if you aren't new to python just skip the steps.

STEP 1: Setting up IDE (Here, PyCharm is being used)

STEP 2: Installing the modules

STEP 3: Setting up your Certificate

STEP 4: Setting up your Excel Sheet

STEP 5: Setting up the position where you want the name to be printed.

STEP 6: Get all your certificates.

## STEP 1: Setting up IDE (Here, PyCharm is being used)
You can use any IDE for running the code - I prefer PyCharm for it's UI and the in-built terminal and consoles.

## STEP 2: Installing the modules
In this code, when initially runned it will result in an error if the following modules are not installed
- Pillow
- pandas

We use Pillow for manipulating our image, while we use pandas to access our data from the excel sheet.
```
pip install Pillow
pip install pandas
```

You can run the command in the Terminal in the folder where the file/code is. Ensure you match the case to not result in an error. If it's already installed you'll get a message like this:
![image](https://user-images.githubusercontent.com/80326865/135320028-f9ff60ac-4ec9-47c5-858a-60fb34026365.png)
You can always check your installed modules in your PyCharm IDE using: File>Settings>Project
![image](https://user-images.githubusercontent.com/80326865/135320587-7d0ce26b-8a86-442f-b609-660cf9ca4198.png)

## STEP 3: Setting up your Certificate
Ensure you add your JPG File into the folder and rename the file name in the code, to match your certificate's name.
![image](https://user-images.githubusercontent.com/80326865/135320886-c184b70d-4076-49a0-aabf-c5102b28901e.png)

## STEP 4: Setting up your Excel Sheet
Ensure you add your Excel File into the folder and rename the file name in the code, to match your file name.
![image](https://user-images.githubusercontent.com/80326865/135321610-da4c24e4-69d2-441c-ba2d-8e0bb3b8dd8a.png)

## STEP 5: Setting up the position where you want the name to be printed.
Open your certificate image in Paint to get the pixel co-ordinates. We need 4 values, x1, x2, y1, y2. This is the text-field dimension where the name will get printed. You can later align it to left, center and right. The coordinates are on the left hand side, bottom corner. 
**NOTE: x1, y1  is left corner while x2, y2 is the right corner.**
![image](https://user-images.githubusercontent.com/80326865/135322281-5caf271e-a8f6-48d1-9b4b-88545cc5c47b.png)

 If it is not in pixels for you, go to the properties page of Paint and change it to pixel:

![image](https://user-images.githubusercontent.com/80326865/135322605-6b12ac70-2896-4204-a7c4-d7b257f703b7.png)

![image](https://user-images.githubusercontent.com/80326865/135322971-dae47a99-2dec-4654-a739-1accef7d3800.png)

Once you have your coordinate of the file go to Line 24 of the code and add your values:

![image](https://user-images.githubusercontent.com/80326865/135323186-3e87b99d-0903-4b67-be9e-e7fa09286f93.png)

As you can see you can also change the values of the font size, and the font style. It will take a few tries for you to get the alignment you want. You can also change it's alignment of text here:

![image](https://user-images.githubusercontent.com/80326865/135323476-1a7c54fd-889e-40b5-80dc-4284311a7af0.png)


## STEP 6: Get all your certificates.
Once, you are done with everything simply run the program. The code will automatically save all the certificates to the same folder as the program. 
![image](https://user-images.githubusercontent.com/80326865/135323684-4f0aaa72-2584-493e-ad40-2041b6911ee4.png)

# IMAGE QUALITY
I really suggest using https://cloudconvert.com/ to convert any PDF file of the certificate template to the PNG File - It let's you control the DPI which further ensures your quality is good for the final certificate that is made.

![image](https://user-images.githubusercontent.com/80326865/135324023-0dd767d2-10a6-4b3a-80d5-ec93dae13acd.png)
![image](https://user-images.githubusercontent.com/80326865/135324054-e5e75941-58e9-486b-b296-a973e813a771.png)

# CONCLUSION
Let me know if there are any other issues you face or any code improvement suggestions one may have. Finally, I'd love to add an extenstion to this which automatically email the final certificate to the participants. Message me on any of my social media platforms to get in touch about this. 



