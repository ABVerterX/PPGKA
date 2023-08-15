import pandas
import re
import sqlalchemy


def main():
    majors = dict()  # all universities' majors, map "A001AN" to data
    majorMapping = dict()  #

    dataFramePlan = pandas.read_excel("plan.xlsx")  # data for plan
    dataFrame2023 = pandas.read_excel("2023.xls", skiprows=1)  # the data for 2023, removed title row
    dataFrame2023.drop(columns="Unnamed: 0", inplace=True)  # Remove the first row which is blank
    dataFrame2022 = pandas.read_excel("2022.xls", skiprows=1)  # the data for 2022, removed title row
    dataFrame2022.drop(columns="Unnamed: 0", inplace=True)  # Remove the first row which is blank
    dataFrame2021 = pandas.read_excel("2021.xls", skiprows=1)  # the data for 2021, removed title row

    commentPattern = re.compile(r"\(.+\)")  # RegEx expression object for matching comments (text in brackets)
    runnerPattern = re.compile(r"\(?保?定?武?汉?\)?\((.+)\)")  # RegEx expression object for matching school-runner

    for _, row in dataFramePlan.iterrows():
        if row["批次"] != "本科常规批":
            continue
        else:
            universityCode = str(row["院校代码"])
            majorId = universityCode + str(row["专业代码"])  # a major's unique identity, like "A00116"
            majors[majorId] = {
                "major_description": match(row["专业名称"], commentPattern),  # a comment of the major
                "school_runner": match(row["招生院校"], runnerPattern, 1),  # the school-runner type of the major
                "university_code": universityCode,
                "tuition": row["学费"]
            }

    for _, row in dataFrame2023.iterrows():
        majorCode = str(row["专业"])[:2]  # the major's code, like "04" in "04能源与动力工程"
        majorName = str(row["专业"])[2:]  # the major's name, like "能源与动力工程" in "04能源与动力工程"
        universityCode = str(row["院校"])[:4]
        universityName = str(row["院校"])[4:]
        majorId = universityCode + majorCode
        majorData = majors.get(majorId, None)
        if notNone(majorData):
            majorData["university_name"] = universityName
            majorData["major_name"] = majorName
            majorData["plan_amount_2023"] = row["投档计划数"]
            majorData["lowest_order_2023"] = row["投档最低位次"]
            # For example, map "A001文科试验班类(文科基础类专业)" to major unique id "A00116"
            majorMapping[universityCode + majorName] = majorId
        else:
            print(f"2023：无法在招生计划中找到 {universityName} 的 {majorName} 专业！")

    for _, row in dataFrame2022.iterrows():
        majorName = str(row["专业代号及名称"])[2:]  # the major's name
        universityCode = str(row["院校代号及名称"])[:4]
        universityName = str(row["院校代号及名称"])[4:]
        majorData = majors.get(majorMapping.get(universityCode + majorName))
        if notNone(majorData):
            majorData["plan_amount_2022"] = row["投档计划数"]
            majorData["lowest_order_2022"] = row["投档最低位次"]
        else:
            print(f"2022：无法在2023年投档志愿表中找到 {universityName} 的 {majorName} 专业！")

    for _, row in dataFrame2021.iterrows():
        majorName = str(row["专业代号及名称"])[2:]  # the major's name
        universityCode = str(row["院校代号及名称"])[:4]
        universityName = str(row["院校代号及名称"])[4:]
        majorData = majors.get(majorMapping.get(universityCode + majorName))
        if notNone(majorData):
            majorData["plan_amount_2021"] = row["投档计划数"]
            majorData["lowest_order_2021"] = row["投档最低位次"]
        else:
            print(f"2021：无法在2023年投档志愿表中找到 {universityName} 的 {majorName} 专业！")

    accumulation = []  # Insert all item into this list first
    for v in majors:
        accumulation.append(majors[v])
    # count = len(accumulation)
    dataFrameResults = pandas.DataFrame(accumulation)  # then create a DataFrame of it

    sqlEngine = sqlalchemy.create_engine("mysql+pymysql://root:RootK1912@localhost:3306/gao_kao")  # connect to MySQL DB
    dataFrameResults.to_sql("plan", sqlEngine, if_exists="replace")
    sqlEngine.dispose()  # close the SQL connexion.


def match(string: str, pattern: re.Pattern, group: int = None) -> str:
    """
    Match strings through RegEx expression.
    :param string: original string
    :param pattern: RegEx pattern object
    :param group: the expr group's order. if the value is None, do not use expr pattern.
    :return: matching result
    """
    result = re.search(pattern, str(string))
    if notNone(group):
        return result.group(group) if notNone(result) else ""
    else:
        return result.group() if notNone(result) else ""


def notNone(value) -> bool:
    """
    Check if the value is None
    """
    return value is not None


main()
