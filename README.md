# CreateNewAccountIteration_VBA
This method (written in Visual Basic) creates a new account iteration. This is very helpful when creating several accounts at the same time, that follows a certain naming convention.

## Background Info
In my field of work, I am required to create multiple accounts per day throughout multiple web environments. These accounts have to follow a certain naming convention per environment.
A naming convention for a new account could be:
"node" + last name + environment suffix + #. 

For example, in our Test environment, "nodemccoytest200" would be a proper naming convention.

These accounts had to be accounted for, followed by the passwords, the default languages, and the country of origin. I created an Access database for this reason, as I enjoy the accessibility of forms and tabs.

Now, imagine having to type out essentially the same username, with account details, MULTIPLE. TIMES. A DAY.

I did this, for about 2 months... then I figured I didn't enjoy the feeling of carpal tunnel, so I created a VB script that did the dirty work for me.

## Take a look...

![Account DB](images/AcctImg01.jpeg)
