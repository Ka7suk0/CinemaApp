# CinemaApp
The cinema application enables users to perform various actions, such as creating accounts, purchasing tickets, and checking the movie listings.  Additionally, the application offers administrative features, including managing movie listings and monitoring ticket sales.

This project was developed in August 2019. At that time, I had only covered basic concepts such as variables in class, along with implementing text printing in an Excel sheet upon button click. Despite not having undergone any formal programming course, I embarked on this project, lacking understanding of loops and functions. My approach was largely trial-and-error, supplemented by extensive research via online resources. Ultimately, though unsure of its inner workings, the project proved to be functional.


Functionality and Specifications:

- User Authentication: Includes account creation (ensuring uniqueness of usernames/emails, validating email format, password verification, security question), login (tracks failed login attempts), password change (verification required), password recovery (via security question), account deletion. Adults have the option to add debit/credit cards for payment. Users can view purchased tickets (pending, paid, used) and generate barcode for cinema entry.

- Administrator Authentication: Allows updating movie listings, accessing ticket sales summary with filters, viewing user list, modifying user details.

- Secret Manager Login: Grants access to password change for administration.

- Movie Listings: Displays movies along with availability dates and ticket purchase options.

- Ticket Purchase: Shows available branches, seats, prices, and timings (which vary during the week and for movie premieres). Payment can be made using registered cards or at the counter.

- Counter: Receives the user-generated ticket barcode and indicates one of three possibilities: unpaid ticket (prompting for payment), already paid ticket.

- Reader (at cinema entrance): Receives the user-generated ticket barcode and indicates one of four possibilities: unpaid ticket, ticket not valid for this branch, paid ticket (and marks it as used), already used ticket.


I plan to resume this project soon, review the code, and optimize it. I'll likely carry out the improvements using Java.

**IMPORTANT:** Enable macros to run the program.
