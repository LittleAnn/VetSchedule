# VetSchedule
A program to create schedules based on disposition data sent by employees.
It takes into account:
1. preferences for work types - day, night, etc.
2. Possibility to enter predetermined days, e.g. if someone has to work continuously or takes a vacation
3. Script is fair in a limited way
4. Script automatically marks weekends based on the selected month
5. Takes into account employee preferences - if someone prefers to work e.g. only day versus night the script should assign preferred shifts to that employee first
6. You can set the maximum number of assigned shifts (i.e. the number of hours)

In the zip there is an excel with random data.

Instructions for using the sheet:
1. dispositions:
* the first column is reserved for consecutive days of the month
* the first row is for employee names
* if an employee has made instructions that he/she can work on a particular day, in the corresponding cell enter 1. If the employee requests time off on that day enter 0

2.Shift limit
* the first column is for employee names
* in the corresponding cell enter the number of shifts of a particular type that you want to assign the maximum in a given month

3. Shift preference
* the first column is for employee names
* in the corresponding cell enter 0 if the employee does not want to work the given type of shifts, if there is such a possibility

4. Fixed shifts
* if we want to assign a predetermined shift to any employee, we should enter the day (without marking the month), script should recognize them as free. In the future versions those should be named something special, making it easier to tell the difference between "free and not asigned" and "free and not available"

