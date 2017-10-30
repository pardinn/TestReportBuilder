# TestReport Builder

This is a simple script to assist in generating test reports for manual testing.
If you don't have any tool to run your tests, take screenshots and generate your reports, this might be helpful.

All you need to do is download the files anywhere you want on your machine and double-click the **Run.vbs** file.

The script takes screenshots in a loop, until you are ready to finish the execution.
It gives you 5 seconds before taking each screenshot, so you have enough time to perform a mouse over, for example.

If you'd like to provide a description in red, you can add the reserved keyword **red_** at the beginning of your description, e.g.: red_This screenshot shows a failure.
This will format that description color in red.

At the end of the execution, you can tell if whether the test is passed of failed, so the script can save your file in the proper folder.
