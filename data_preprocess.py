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

    data.sort_values(by=['课程名称', '课程类型', '任课教师'], inplace=True)

    # 处理有老师只教授一门课，但是却有多条记录的情况
    # 找出每个教师的所有课程记录
    teacher_courses = data.groupby('任课教师').agg({
        '课程名称': lambda x: set(x),
        '课程类型': lambda x: set(x)
    }).reset_index()

    # 找出所有记录课程名称和类型都相同的教师
    same_course_teachers = teacher_courses[
        (teacher_courses['课程名称'].map(len) == 1) & 
        (teacher_courses['课程类型'].map(len) == 1)
    ]['任课教师']

    # 更新这些教师的情况说明
    for teacher in same_course_teachers:
        mask = data['任课教师'] == teacher
        if data.loc[mask, '情况说明'].eq('任选其一').all():
            data.loc[mask, '情况说明'] = '在该老师这门课程中任选其一'
            
    # 处理所有课程都是任选其一的老师
    # 找出所有课程都是"任选其一"的教师
    all_optional_teachers = data.groupby('任课教师').filter(
        lambda x: x['情况说明'].eq('任选其一').all()
    )['任课教师'].unique()

    potential_deletions = []
    teachers_to_update = []

    for teacher in all_optional_teachers:
        teacher_courses = data[data['任课教师'] == teacher]
        
        # 如果教师有不同的课程
        if len(teacher_courses[['课程名称', '课程类型']].drop_duplicates()) > 1:
            # 检查该教师的所有课程是否都有其他教师教授
            all_courses_shared = True
            for _, course_row in teacher_courses.iterrows():
                other_teachers = data[
                    (data['课程名称'] == course_row['课程名称']) &
                    (data['课程类型'] == course_row['课程类型']) &
                    (data['任课教师'] != teacher)
                ]
                if len(other_teachers) == 0:
                    all_courses_shared = False
                    break
            
            if all_courses_shared:
                # 找出记录数最多的课程
                course_counts = teacher_courses.groupby(['课程名称', '课程类型']).size()
                is_single_record = course_counts.max() == 1
                max_count_course = course_counts.sample(1).index[0] if is_single_record else course_counts.idxmax()
                
                # 存储需要更新的教师和课程信息，同时记录是否为单条记录
                teachers_to_update.append({
                    '任课教师': teacher,
                    '课程名称': max_count_course[0],
                    '课程类型': max_count_course[1],
                    'is_single': is_single_record
                })
                
                # 存储可能需要删除的记录
                for _, course_row in teacher_courses.iterrows():
                    if (course_row['课程名称'], course_row['课程类型']) != max_count_course:
                        potential_deletions.append({
                            '任课教师': teacher,
                            '课程名称': course_row['课程名称'],
                            '课程类型': course_row['课程类型']
                        })

    # 验证删除操作是否全
    safe_deletions = []
    for deletion in potential_deletions:
        # 检查删除后是否还有其他教师教授该课程
        remaining_records = data[
            (data['课程名称'] == deletion['课程名称']) &
            (data['课程类型'] == deletion['课程类型']) &
            (~data['任课教师'].isin([d['任课教师'] for d in potential_deletions if 
                d['课程名称'] == deletion['课程名称'] and 
                d['课程类型'] == deletion['课程类型']]))
        ]
        if len(remaining_records) > 0:
            safe_deletions.append(deletion)

    # 执行安全的删除操作
    for deletion in safe_deletions:
        mask = (
            (data['任课教师'] == deletion['任课教师']) &
            (data['课程名称'] == deletion['课程名称']) &
            (data['课程类型'] == deletion['课程类型'])
        )
        data = data[~mask]

    # 更新选定课程的情况说明
    for update in teachers_to_update:
        mask = (
            (data['任课教师'] == update['任课教师']) &
            (data['课程名称'] == update['课程名称']) &
            (data['课程类型'] == update['课程类型'])
        )
        # 如果是单条记录，设置为"必选"，否则设置为"必须选定该老师这门课程的任选其一"
        data.loc[mask, '情况说明'] = '必选' if update['is_single'] else '在该老师这门课程中任选其一'

    # 新增处理逻辑
    # 找出所有"任选其一"相关的记录并按课程分组
    optional_courses = data[
        (data['情况说明'].isin(['任选其一', '在该老师这门课程中任选其一'])) &
        (data.groupby(['课程名称', '课程类型'])['任课教师'].transform('count') > 1)
    ]
    
    course_groups = optional_courses.groupby(['课程名称', '课程类型'])
    
    records_to_delete = []
    
    for (course_name, course_type), group in course_groups:
        has_teacher_specific = '在该老师这门课程中任选其一' in group['情况说明'].values
        
        if has_teacher_specific:
            # 处理情况1：存在"在该老师这门课程中任选其一"的组别
            optional_teachers = group[group['情况说明'] == '任选其一']['任课教师'].unique()
            if len(optional_teachers) == 0:
                continue
                
            for teacher in optional_teachers:
                # 检查该教师的其他课程
                other_courses = data[data['任课教师'] == teacher]
                other_courses = other_courses[
                    ~((other_courses['课程名称'] == course_name) & 
                      (other_courses['课程类型'] == course_type))
                ]
                
                # 检查是否有重要课程
                has_important_course = False
                
                # 检查必选或特定课程
                if other_courses['情况说明'].isin(['必选', '在该老师这门课程中任选其一']).any():
                    has_important_course = True
                
                # 检查是否有独教的任选其一课程
                for _, other_course in other_courses[other_courses['情况说明'] == '任选其一'].iterrows():
                    course_teachers = data[
                        (data['课程名称'] == other_course['课程名称']) &
                        (data['课程类型'] == other_course['课程类型'])
                    ]['任课教师'].unique()
                    if len(course_teachers) == 1:
                        has_important_course = True
                        break
                
                if has_important_course:
                    records_to_delete.append({
                        '任课教师': teacher,
                        '课程名称': course_name,
                        '课程类型': course_type
                    })
        else:
            # 处理情况2：只有"任选其一"的组别
            potential_deletions = []
            for _, row in group.iterrows():
                teacher = row['任课教师']
                # 检查该教师的其他课程
                other_courses = data[data['任课教师'] == teacher]
                other_courses = other_courses[
                    ~((other_courses['课程名称'] == course_name) & 
                      (other_courses['课程类型'] == course_type))
                ]
                
                # 检查是否有重要课程
                has_important_course = False
                
                # 检查必选或特定课程
                if other_courses['情况说明'].isin(['必选', '在该老师这门课程中任选其一']).any():
                    has_important_course = True
                
                # 检查是否有独教的任选其一课程
                for _, other_course in other_courses[other_courses['情况说明'] == '任选其一'].iterrows():
                    course_teachers = data[
                        (data['课程名称'] == other_course['课程名称']) &
                        (data['课程类型'] == other_course['课程类型'])
                    ]['任课教师'].unique()
                    if len(course_teachers) == 1:
                        has_important_course = True
                        break
                
                if has_important_course:
                    potential_deletions.append({
                        '任课教师': teacher,
                        '课程名称': course_name,
                        '课程类型': course_type
                    })
            
            # 检查删除后是否还有其他教师
            remaining_teachers = set(group['任课教师']) - {d['任课教师'] for d in potential_deletions}
            if remaining_teachers:  # 如果删除后还有教师剩余
                records_to_delete.extend(potential_deletions)
    
    # 执行删除操作
    for deletion in records_to_delete:
        mask = (
            (data['任课教师'] == deletion['任课教师']) &
            (data['课程名称'] == deletion['课程名称']) &
            (data['课程类型'] == deletion['课程类型'])
        )
        data = data[~mask]
    
    # 更新剩余单条记录的情况说明
    single_records = data.groupby(['课程名称', '课程类型']).filter(lambda x: len(x) == 1)
    data.loc[single_records.index, '情况说明'] = '必选'
    
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
    # 匹配督导姓名与课程数据
    supervisor_course_data = filter_data[filter_data['任课教师'].isin(supervisor_data['姓名'])]
    # 对筛选得到的数据按"任课教师"分组，并对"开课校区"和"课表信息"列进行聚合处理，在聚合过程种，"开课校区"列的值会去重并用逗号连接，"课表信息"列的值会直接用逗号连接
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