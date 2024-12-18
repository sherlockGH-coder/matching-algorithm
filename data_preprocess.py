import pandas as pd


def read_data(file_path):
    """
    读取Excel文件数据。

    Args:
        file_path (str): Excel文件的路径。

    Returns:
        pandas.DataFrame: 包含Excel文件数据的DataFrame。
    """
    return pd.read_excel(file_path)


def filter_course_data(course_data):
    """
    过滤课程数据，移除特定类型的课程。

    Args:
        course_data (pandas.DataFrame): 原始课程数据。

    Returns:
        pandas.DataFrame: 过滤后的课程数据。

    Notes:
        过滤条件包括：
        - 删除设计(论文)课程
        - 删除开课校区为揭阳校区的课程
        - 删除计划类型为重修计划的课程
        - 删除教学班名称包含"补修"的记录
    """
    course_data = course_data[course_data['课程类别'] != '设计(论文)'] 
    course_data = course_data[course_data['开课校区'] != '揭阳校区']
    course_data = course_data[course_data['计划类型'] != '重修计划']
    course_data = course_data[~course_data['教学班名称'].str.contains('补修')]
    return course_data


def add_course_type(data):
    """
    添加课程类型并处理数据。

    Args:
        data (pandas.DataFrame): 原始课程数据。

    Returns:
        pandas.DataFrame: 处理后的课程数据。

    Notes:
        - 拆分多个教师的记录
        - 添加课程类型（实验/理论）
        - 复制理论和实践学时都大于0的记录
        - 添加情况说明（必选/任选其一）
    """
    data['任课教师'] = data['任课教师'].apply(lambda x: x.split(','))
    data = data.explode('任课教师')
    data['课程类型'] = data.apply(lambda x: '实验' if (x['实践学时'] > 0) & (x['理论学时'] == 0 or pd.isna(x['理论学时'])) else '理论', axis=1)
    data_experiment = data[(data['理论学时'] > 0) & (data['实践学时'] > 0)].copy()
    data_experiment['课程类型'] = '实验'
    data = pd.concat([data, data_experiment])
    data = data.sort_values(by=['课程名称', '课程类型'])
    data['情况说明'] = data.apply(lambda x: '必选' if ((data['课程名称'] == x['课程名称']) & (data['课程类型'] == x['课程类型'])).sum() == 1 or (data['任课教师'] == x['任课教师']).sum() == 1 else '任选其一', axis=1)
    return data


def get_supervisor_course_data(supervisor_data, filter_data):
    """
    获取督导课程信息。

    Args:
        supervisor_data (pandas.DataFrame): 督导信息数据。
        filter_data (pandas.DataFrame): 过滤后的课程数据。

    Returns:
        pandas.DataFrame: 督导课程信息。

    Notes:
        - 匹配督导姓名与课程数据
        - 合并开课校区和课表信息
        - 为没有课程的督导添加空记录
    """
    supervisor_course_data = filter_data[filter_data['任课教师'].isin(supervisor_data['姓名'])]
    supervisor_course_data = supervisor_course_data.groupby('任课教师').agg({'开课校区': lambda x: ','.join(set(x.dropna().astype(str))), '课表信息': lambda x: ','.join(x.dropna().astype(str))}).reset_index()
    
    for index, row in supervisor_data.iterrows():
        if row['姓名'] not in supervisor_course_data['任课教师'].values:
            new_row = {'任课教师': row['姓名'], '开课校区': '', '课表信息': '{}'}
            supervisor_course_data = pd.concat([supervisor_course_data, pd.DataFrame([new_row])], ignore_index=True)

    return supervisor_course_data


def get_task_data(filter_data):
    """
    获取听课任务数据。

    Args:
        filter_data (pandas.DataFrame): 过滤后的课程数据。

    Returns:
        pandas.DataFrame: 听课任务数据。

    Notes:
        按课程名称、课程类型和情况说明分组，合并其他列的数据。
    """
    filter_data = filter_data.groupby(['课程名称', '课程类型', '情况说明']).agg(lambda x: ','.join(set(x.dropna().astype(str)))).reset_index()
    return filter_data


def get_data(course_file, supervisor_file):
    """
    处理并获取督导课程数据和听课任务数据。

    Args:
        course_file_path (str): 课程数据文件路径。
        supervisor_file_path (str): 督导数据文件路径。

    Returns:
        tuple: 包含督导课程数据和听课任务数据的元组。
            - supervisor_course_data (pandas.DataFrame): 督导课程数据。
            - task_data (pandas.DataFrame): 听课任务数据。
    """
    # 判断course_file_path和supervisor_file_path是否为字符串
    if isinstance(course_file, str) and isinstance(supervisor_file, str):
        course_data = read_data(course_file)
        supervisor_data = read_data(supervisor_file)
    # 判断是否为pandas.DataFrame
    elif isinstance(course_file, pd.DataFrame) and isinstance(supervisor_file, pd.DataFrame):
        course_data = course_file
        supervisor_data = supervisor_file
    # 抛出异常
    else:
        raise ValueError('输入数据类型错误')
    
    filter_data = filter_course_data(course_data)
    filter_data = add_course_type(filter_data)
    supervisor_course_data = get_supervisor_course_data(supervisor_data, filter_data)
    task_data = get_task_data(filter_data)
    
    supervisor_index_map = {name: idx for idx, name in enumerate(supervisor_course_data['任课教师'])}
    task_index_map = {tuple(row): idx for idx, row in task_data[['课程名称', '课程类型', '情况说明']].iterrows()}

    # 添加索引列
    supervisor_course_data['索引'] = supervisor_course_data['任课教师'].map(supervisor_index_map)
    task_data['索引'] = task_data.apply(lambda row: task_index_map[(row['课程名称'], row['课程类型'], row['情况说明'])], axis=1)

    return supervisor_course_data, task_data, filter_data


if __name__ == "__main__":
    course_file_path = './2024春开课任务.xls'
    supervisor_file_path = './督导名单.xlsx'
    supervisor_course_data, task_data, fliter_data = get_data(course_file_path, supervisor_file_path)
    print(supervisor_course_data.head(1))
    