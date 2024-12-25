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
        - 删除特定条件下的重复课程记录
    """
    # 将任课教师列中的多个教师记录拆分成多行
    data['任课教师'] = data['任课教师'].apply(lambda x: x.split(','))
    data = data.explode('任课教师')
    
    # 处理课表信息，只保留对应教师的课表
    data['课表信息'] = data.apply(lambda row:
        ','.join(['{' + schedule + '}' for schedule in str(row['课表信息']).strip('{}').split('},{')
                 if row['任课教师'] in schedule.replace('[主讲]', '')])
        if not pd.isna(row['课表信息']) else '{}',
        axis=1
    )
    
    # 根据实践学时和理论学时添加并判断课程类型
    data['课程类型'] = data.apply(lambda x: '实验' if (x['实践学时'] > 0) & (x['理论学时'] == 0 or pd.isna(x['理论学时'])) else '理论', axis=1)
    
    # 将同时有理论学时和实践学时的课程复制一份
    data_experiment = data[(data['理论学时'] > 0) & (data['实践学时'] > 0)].copy()
    data_experiment['课程类型'] = '实验'
    data = pd.concat([data, data_experiment])
    
    return data


def process_course_data(data):
    # 按课程名称和课程类型排序
    data = data.sort_values(by=['课程名称', '课程类型'])
    
    # 按课程维度标记情况说明
    data['情况说明'] = data.apply(
        lambda x: '必选' if len(data[
            (data['课程名称'] == x['课程名称']) & 
            (data['课程类型'] == x['课程类型'])
        ]) == 1 else '任选其一',
        axis=1
    )

    # 再按任课教师维度更新情况说明
    data.loc[data['任课教师'].map(lambda x: len(data[data['任课教师'] == x])) == 1, '情况说明'] = '必选'
    
    # 找出所有"任选其一"的课程组
    optional_courses = data[data['情况说明'] == '任选其一']
    course_groups = optional_courses.groupby(['课程名称', '课程类型'])
    
    # 存储需要删除的记录
    records_to_remove = []
    
    # 处理每个课程组
    for (course_name, course_type), group in course_groups:
        teachers = group['任课教师'].unique()
        teachers_with_mandatory = []
        
        # 检查每个教师是否有必选课程
        for teacher in teachers:
            has_mandatory = data[
                (data['任课教师'] == teacher) & 
                (data['情况说明'] == '必选')
            ].shape[0] > 0
            
            if has_mandatory:
                teachers_with_mandatory.append(teacher)
        
        # 如果不是所有教师都有必选课程，则处理删除逻辑
        if teachers_with_mandatory and len(teachers_with_mandatory) < len(teachers):
            for teacher in teachers_with_mandatory:
                records_to_remove.append({
                    '任课教师': teacher,
                    '课程名称': course_name,
                    '课程类型': course_type
                })
    
    # 删除标记的记录
    if records_to_remove:
        for record in records_to_remove:
            mask = (
                (data['任课教师'] == record['任课教师']) & 
                (data['课程名称'] == record['课程名称']) & 
                (data['课程类型'] == record['课程类型'])
            )
            data = data[~mask]
    
    # 更新剩余单条记录的状态
    course_counts = data.groupby(['课程名称', '课程类型']).size()
    single_courses = course_counts[course_counts == 1].index
    
    for course_name, course_type in single_courses:
        mask = (
            (data['课程名称'] == course_name) & 
            (data['课程类型'] == course_type)
        )
        data.loc[mask, '情况说明'] = '必选'
    
    return data


def get_supervisor_course_data(supervisor_data, course_data):
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
    # 匹配督导姓名与课程数据
    supervisor_course_data = course_data[course_data['任课教师'].isin(supervisor_data['姓名'])]
    # 对筛选得到的数据按“任课教师”分组，并对“开课校区”和“课表信息”列进行聚合处理，在聚合过程种，“开课校区”列的值会去重并用逗号连接，“课表信息”列的值会直接用逗号连接
    supervisor_course_data = supervisor_course_data.groupby('任课教师').agg({'开课校区': lambda x: ','.join(set(x.dropna().astype(str))), '课表信息': lambda x: ','.join(x.dropna().astype(str))}).reset_index()
    # 函数遍历supervisor_data中的每一行，检查每个督导是否在 supervisor_course_data中。如果某个督导不在supervisor_course_data 中，函数会为该督导添加一条空记录，记录中开课校区和课表信息列为空字符串或空字典。
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
    
    # 过滤课程数据，移除不需要匹配的课程任务
    filter_data = filter_course_data(course_data)
    # 添加课程类型并处理数据
    filter_data = add_course_type(filter_data)
    # 处理课程数据
    filter_data = process_course_data(filter_data)
    # 获取督导课程信息
    supervisor_course_data = get_supervisor_course_data(supervisor_data, course_data)
    # 获取听课任务数据
    task_data = get_task_data(filter_data)
    
    # 创建索引映射
    supervisor_index_map = {name: idx for idx, name in enumerate(supervisor_course_data['任课教师'])}
    task_index_map = {tuple(row): idx for idx, row in task_data[['课程名称', '课程类型', '情况说明']].iterrows()}

    # 添加索引列
    supervisor_course_data['索引'] = supervisor_course_data['任课教师'].map(supervisor_index_map)
    task_data['索引'] = task_data.apply(lambda row: task_index_map[(row['课程名称'], row['课程类型'], row['情况说明'])], axis=1)

    return supervisor_course_data, task_data, filter_data


if __name__ == "__main__":
    course_file_path = '../test-algorithm/2024春开课任务.xls'
    supervisor_file_path = '../test-algorithm/督导名单.xlsx'
    supervisor_course_data, task_data, filter_data = get_data(course_file_path, supervisor_file_path)
    print(supervisor_course_data.head(1))
    supervisor_course_data.to_excel('../test-algorithm/supervisor_course_data.xlsx', index=False)
    task_data.to_excel('../test-algorithm/task_data.xlsx', index=False)
    filter_data.to_excel('../test-algorithm/filter_data.xlsx', index=False)