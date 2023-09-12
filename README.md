# schedule_checker
This is a schedule checking tool written in Excel / VBA that can evaluate across multiple graduation requirements.

# background
this is a version of a tool I use in my work to help students make study-abroad plans and move the courses around. 

# this version / instructions
I've stripped out the actual data for students and instructors to give a model case. Basically, by using the names of courses in different tabs, it creates a part of a requirement.

E.g. In
####
履修基準分野	学年	授業数	履修授業数	状況	科目名称
教養科目（共通基礎科目・日本国憲法）	2	1	1	OKAY	Civics
教養科目（現代的教養科目）	1	2	2	OKAY	French I	French II	Farming I	Farming II	Spanish I	Spanish II	German I	German II	School Bullying	Intro to Anthropology	Music Survey	Art Survey	Art Survey	Contemporary Literature Survery	Chinese I	Chinese II	Korean I	Korean II	Drama	Soul Studies	Finance	Life: An Introduction	General Humanities	Anthropology I	Anthropology II	International Understanding																
####

The second line will tell the system to check if the student is taking "Civics". If the answer is "yes" then the line is OKAY
The third line will tell the system to check that the student is taking at least two calsses from that section. if they have zero or 1 NG, 2 or more OKAY

If all rows are OKAY, then the section will return OKAY on the main schedule

Students enter courses by entering the course number in the `1Sheet` sheet. This will automatically:
(1) generate a calendar for that semester (VBA)
(2) highlight whether they are in the wrong type of semester (spring vs fall)
(3) update the requirements based on added/deleted courses

The system can also adapt to different lists for different majors.


The courses are populated in `授業リスト` and as long as they are well-formatted there's no limit to the courses. Columns N-W in the course list are used to speed up the checks by enabling VBA to use dictionaries (a dictionary in this context is basically a pre-made list).
