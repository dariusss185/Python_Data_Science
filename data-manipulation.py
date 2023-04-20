import pandas as pd

df = pd.read_csv("MOCK_DATA.csv")


#print(df.head(5)) #print top=head bottom-tail  5 lines
male_person= df[(df['gender'] == 'Male')]
print(male_person.head(5))
print(len(male_person), "total males" )#total

female_person= df[(df['gender'] == 'Female')]
print(female_person.head(5))
print(len(female_person), "total females")#total


female_person.to_csv('females.csv')
