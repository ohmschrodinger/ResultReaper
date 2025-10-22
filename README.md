# ResultReaper

ResultReaper is a web automation tool that fetches student results from [https://siuexam.siu.edu.in/forms/resultview.html](https://siuexam.siu.edu.in/forms/resultview.html) and saves them locally on your device.

You simply enter the range of seat numbers and PRNs, and the tool handles the rest. Seat numbers and PRNs are compared, mapped, and saved in `mappings.csv`.

Additionally, there is a script for Result Analysis which provides insights by extracting data from the saved result PDFs. These reports include:

* overall GPA statistics,
* grade distributions,
* subject wise performance analysis,
* and correlation studies between subjects and gpa.

These help identify anomalies, such as if a particular division received significantly lower grades than others or if a student performs well in labs but poorly in the corresponding theory exams.

Correlation studies help you understand which subjects contribute more positively or negatively to your GPA. The analysis also calculates the average grade of all students in a particular subject compared to your grade.

The report includes variance and skewness in GPA statistics to help determine whether the grade distribution is symmetric or skewed.

# Why is this important?

Because the uni uses a percentile based grading system, where fixed percentages of students get each grade like the top 3% get a 10, the next 12% get a 9, and so on.

This works okay if the exam scores are spread out, so the top 3% really reflects the highest scorers. But in a lot of cases, the assignments such as poster making or report writing many students get very similar or even perfect marks because of how the grading works. Even though there’s a rubric, the final marks can feel random or based on luck. 

For example, if out of 100 students, 10 students all get 20/20, all 10 get a 10 grade, which is fair. But then, the next students who scored 19/20 end up getting a 9 grade, just because they fall outside the top 3%. the system doesn’t consider how close their marks actually are (like through z-scores or any statistical measure) is just sticks to the percentiles.

So, many students who scored nearly the same marks get different grades, which feels unfair.

(and yes sql injection would make fetching the results easier)
