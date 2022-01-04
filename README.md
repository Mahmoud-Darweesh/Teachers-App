# Teachers-App
Allows teachers to award students who excel at their work.

Teacher’s App
The Introduction:

The students performance app was an idea my teacher gave me in order to help teachers evaluate students' participation in class and subsequently award them for their efforts. 

The Database:

I wasn’t sure where to start in the beginning but I decided to begin with finding a good place to store all the data. I had many options but I wanted it to be simple since its my last year in the school and I want this app to be usable even when I leave. So after weighing my options I found a great solution which was google sheets. Although I did not know that large amounts of requests could be done through google's api to 1 google sheet was possible I was excited, it was the perfect place to store the information. It was easily accessible to any one even if they are not a tech savvy person, it was reliable so I did not need to worry about it shutting down, and lastly I was able to send and receive information with little to no delay. 

The Prototype:

After deciding on the place I was going to store all the data in I began work on a small prototype that did the main function of the application which was to award students different types of points based on their work. I wanted to make it easy for the teachers to click the buttons so that they won't waste time by swiping their finger across the screen from the students name to the button. So I added the ability to click the side of the name and the whole row would light up blue. Then I decided that the different categories of work that a student did would correspond to a star and from there my first prototype was done.

![Prototype](https://user-images.githubusercontent.com/95291720/147952922-db5bc053-71e9-43dd-a446-b379edc70d02.png)

The Admin Tab[^1]:

Next I needed a way for the head of department to view the work of the students and teachers in a single place. So I created the admin tab[^1], I knew that in a real scenario the admins won’t want to see every student and every class all at once so I made it easy for the user to choose what exactly they wanted to see by choosing between 2 categories classes and students. Then I added an advanced feature that allows the user to sort by Date Grade Name and Teacher.  After that I needed to make the information printable in order for the admins to have a real world copy of what they saw on screen. So I added a print function that takes the data-grid and sizes it down to the size of an a4 paper and adds the school logo and other basic information.

  Loading All Information Without Sorting:

![Admin Controls 1](https://user-images.githubusercontent.com/95291720/147954159-f92da92b-7c00-446c-ac25-8cf652ca629f.PNG)
  
  Sorting By Date and Grade:
  
  ![Admin Controls 2](https://user-images.githubusercontent.com/95291720/147954257-465ee975-5b78-489b-8352-70e8bd1bcc03.PNG)
  
   The Printed report for it:
   
   [Custom Report.pdf](https://github.com/Mahmoud-Darweesh/Teachers-App/files/7802980/Custom.Report.pdf)

  Sorting By Name:
  
  ![Admin Controls 3](https://user-images.githubusercontent.com/95291720/147954297-52519fdf-04ca-49f0-8c1d-44be6188db5f.PNG)

  Basic Sorting:
  
  ![Admin Controls 4](https://user-images.githubusercontent.com/95291720/147954364-90d60e38-f32a-4986-b990-ed0c2286e747.PNG)
  
  ![Admin Controls 5](https://user-images.githubusercontent.com/95291720/147954390-76afeaa6-279f-4f9d-a899-7800ec9de8d5.PNG)
  
  The Printed Report For It:
  
  [11 PreCalculus's Report.pdf](https://github.com/Mahmoud-Darweesh/Teachers-App/files/7802982/11.PreCalculus.s.Report.pdf)
  
The Loading screen:

After I was done with the admin page[^1] I realised there was an issue I needed to solve and that was the loading times. Since I was only using a small dataset when I was testing the application I never ran into any issue until it came time to import the real information. The admin page[^1] was taking long amounts of time to load after every action, I kept trying different solutions like loading the data in chunks but none of them were practical. So I decided I would need to change the way the app dealt with information. I began by seeing how other applications went about doing it and found the answer I was looking for. I needed to add a loading screen where all the information will be loaded once then stored in a public array for the rest of the classes to use in their computation. This was a massive success. I was able to remove the loading time completely and provide the user with a much more enjoyable experience.
  
Previous version:

  Slow Load Times:
  
https://user-images.githubusercontent.com/95291720/147951760-7004d567-b2f4-4645-bf2b-44029fdd1979.mp4

  Previous Implementation:
  
  ![Slow Load of Data](https://user-images.githubusercontent.com/95291720/147952207-42307ab3-d0ad-4d7e-942a-03014f2bd33f.PNG)
  
  ![Slow Load of Data 2](https://user-images.githubusercontent.com/95291720/147952219-ab7c37ba-eb5c-4586-b66a-658a7203148d.PNG)

Newer Version:

  Loading Screen:
  
  ![Splash Screen](https://user-images.githubusercontent.com/95291720/147955684-67222076-724f-4235-a61d-8d8921ddd7c0.PNG)

  Public variables:
  
  ![Fast Load Of Data Splash Screen Code 3](https://user-images.githubusercontent.com/95291720/147951030-90e1dcba-63de-46e6-90fc-50768704b5dd.PNG)
  
  Using methods to load the information:
  
  ![Fast Load Of Data Splash Screen Code 2](https://user-images.githubusercontent.com/95291720/147950941-8f9e90ef-5728-4f34-8614-349504433d68.PNG)
  
  Example of the method:
  
  ![Fast Load Of Data Splash Screen Code](https://user-images.githubusercontent.com/95291720/147951050-1afade88-7ec8-4f38-a843-b27679e2070c.PNG)



The Sign In Page:

To differentiate between teacher accounts and admin accounts I needed to create a sign in page, so that if they were a teacher account they would have access to the performance tab[^2] and to the awards tab[^4], and if they were an admin they would have 2 additional tabs: the admin tab[^1] and the notifications tab[^3].
  
  Sign In Page:
  
  ![Sign In](https://user-images.githubusercontent.com/95291720/147951904-26d0aba0-3fed-43e9-b33b-4fc4c957006f.PNG)
  
  Teachers Dashboard:
  
  ![Teacher Main Dashboard](https://user-images.githubusercontent.com/95291720/147951952-2ecae533-477f-4201-99f9-a90bd38e46ad.PNG)
  
  Admins Dashboard:
  
  ![Main Dashboard](https://user-images.githubusercontent.com/95291720/147952262-ca113cfb-82b1-4232-ba24-589fd1350571.PNG)


Publishing the app:

In the end I found it difficult to install the application on the devices of the teachers because I had the application in the form of an executable file which was not easy to send especially when there are multiple libraries that need to be sent with it. So I decided to learn how to make .msi files and after a short while I was able to condense and publish the app to all the teachers safely and in a timely manner.

![Publish](https://user-images.githubusercontent.com/95291720/147955813-5067a7a3-32b6-468b-888d-1e9cb7cf8c15.PNG)

The tabs:

In the performance tab you can choose the class you want from the dropdown menu then after you have chosen an option you will see the grid view being populated with the names and scores of all the students in that class. The buttons on the right side allow the user to increase that specific score by one and to avoid any misscliks a popup saying that “the student was awarded a star” is shown.

  Finding a Class:
  
  ![Preformance Tab 1](https://user-images.githubusercontent.com/95291720/147957977-c9725929-9a8e-4fa3-bab6-34b56a181850.PNG)
  
  The Class:
  
  ![Preformance Tab 2](https://user-images.githubusercontent.com/95291720/147958009-1852eeb4-009a-4e45-b871-6d7d55e5dcd7.PNG)

In the awards tab you can choose a class and print the certificates for the students in that class by double clicking on that specific student and his achievement.

  Awards For a Class:
  
  ![Awards](https://user-images.githubusercontent.com/95291720/147956216-7ba5493c-390c-4d1c-86ba-0b7359655bd4.PNG)
  
  Report For That Class:
  
  [Awards 11 PreCalculus Report.pdf](https://github.com/Mahmoud-Darweesh/Teachers-App/files/7803016/Awards.11.PreCalculus.Report.pdf)

  Certificate for a student:
  
  [AMRO's Master Mathematician Certificate.pdf](https://github.com/Mahmoud-Darweesh/Teachers-App/files/7803007/AMRO.s.Master.Mathematician.Certificate.pdf)
  
In the admin tab you can view the work of all the students and teachers throughout the year and sort them by any of the given categories which are Basic: classes, students and Advanced: Date,Grade,Student and Teacher.

In the notification tab an admin can see when a student has achieved a new certificate there they can print that certificate and in doing so the notification is removed.

  Example of a Notification:
  
  ![Notifications](https://user-images.githubusercontent.com/95291720/147956445-7ae974c5-db17-4910-a629-cb55c6fc0350.PNG)
  
# In Progress:
  
  Right now I am trying to and a new tab that can change basied on the time of use. So if it is during school hours it would show the class that the teacher has automatically and if it is after school it would show the certificates that need to be printed.

## References:

[^1]: In the admin tab you can view the work of all the students and teachers throughout the year and sort them by any of the given categories which are Basic: classes, students and Advanced: Date,Grade,Student and Teacher.

[^2]: In the performance tab you can choose the class you want from the dropdown menu then after you have chosen an option you will see the grid view being populated with the names and scores of all the students in that class. The buttons on the right side allow the user to increase that specific score by one and to avoid any misscliks a popup saying that “the student was awarded a star” is shown.

[^3]: In the notification tab an admin can see when a student has achieved a new certificate there they can print that certificate and in doing so the notification is removed.

[^4]: In the awards tab you can choose a class and print the certificates for the students in that class by double clicking on that specific student and his achievement.
