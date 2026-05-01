import pandas as pd

data = [
    {
        "Name": "Alice Smith",
        "RollNo": "A102",
        "Class": "Grade 10",
        "Section": "A",
        "Contact": "alice@example.com",
        "FeeAmount": 5000,
        "FeeType": "Tuition",
        "FeeRemarks": "Paid in Full",
        "LateFee": 0
    },
    {
        "Name": "Bob Jones",
        "RollNo": "A103",
        "Class": "Grade 10",
        "Section": "B",
        "Contact": "bob@example.com",
        "FeeAmount": 200,
        "FeeType": "Transport",
        "FeeRemarks": "Bus Fee",
        "LateFee": 50
    }
]

df = pd.DataFrame(data)
df.to_excel("stu.xlsx", index=False)
print("Created stu.xlsx")
