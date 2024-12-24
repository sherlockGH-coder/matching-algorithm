import pandas as pd
import numpy as np
import re

class TimeParser:
    """
    用于解析课程时间信息的类。

    Attributes:
        course_times (list): 包含课程时间信息的列表。
    """

    def __init__(self, raw_course_data, time_column):
        """
        初始化TimeParser对象。

        Args:
            raw_course_data (pandas.DataFrame): 包含课程数据的DataFrame。
            time_column (str): 包含时间信息的列名。
        """
        # 获取课程时间信息列表
        self.course_times = raw_course_data[time_column].tolist()

    def parse_weeks(self, week_str):
        """
        解析周数字符串。

        Args:
            week_str (str): 表示周数的字符串，例如 "5周" 或 "1-5,8-18周"。

        Returns:
            set: 包含所有周数的集合。

        Example:
            >>> parser = TimeParser(df, 'time_column')
            >>> parser.parse_weeks("1-5,8-18周")
            {1, 2, 3, 4, 5, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18}
        """
        week_str = week_str.replace('周', '')
        parts = week_str.split(',')
        weeks = set()
        for part in parts:
            if '-' in part:
                start, end = part.split('-')
                weeks.update(range(int(start), int(end)+1))
            else:
                weeks.add(int(part))
        return weeks

    def parse_periods(self, period_str):
        """
        解析节次字符串。

        Args:
            period_str (str): 表示节次的字符串，例如 "第01,02节" 或 "第06,07,08,09节"。

        Returns:
            set: 包含所有节次的集合。

        Example:
            >>> parser = TimeParser(df, 'time_column')
            >>> parser.parse_periods("第01,02节")
            {1, 2}
        """
        period_str = period_str.replace('第', '').replace('节', '')
        parts = period_str.split(',')
        periods = set()
        for part in parts:
            if '-' in part:
                start, end = part.split('-')
                periods.update(range(int(start), int(end)+1))
            else:
                periods.add(int(part))
        return periods

    def parse_day(self, day_str):
        """
        将星期字符串转换为数字表示。

        Args:
            day_str (str): 表示星期的字符串，例如 "星期一"。

        Returns:
            int: 星期的数字表示，星期一为1，依此类推。如果输入无效，返回0。

        Example:
            >>> parser = TimeParser(df, 'time_column')
            >>> parser.parse_day("星期三")
            3
        """
        days = {
            '星期一':1, '星期二':2, '星期三':3, 
            '星期四':4, '星期五':5, '星期六':6, '星期日':7
        }
        return days.get(day_str, 0)
    
    def extract_course_time(self):
        """
        提取课表信息。

        Returns:
            list: 包含每个课程时间信息的列表。

        Example:
            >>> parser = TimeParser(df, 'time_column')
            >>> parser.extract_course_time()
            [['{...}', '{...}'], [], [{...}'], ...]
        """
        schedules = []
        for course_time in self.course_times:
            if isinstance(course_time, str):
                schedule_pattern = r'\{.*?\}'
                schedule = re.findall(schedule_pattern, course_time)
                schedules.append(schedule)
            else:
                schedules.append([])
        return schedules

    def parse_class_entry(self, entry_str):
        """
        解析单个课表条目。

        Args:
            entry_str (str): 表示单个课表条目的字符串，例如 "[刘高勇],[星期一],[第01,02节],[5周],[实A南-705]"。

        Returns:
            dict: 包含解析后信息的字典，如果解析失败则返回None。

        Example:
            >>> parser = TimeParser(df, 'time_column')
            >>> parser.parse_class_entry("[刘高勇],[星期一],[第01,02节],[5周],[实A南-705]")
            {'teacher': '刘高勇', 'day': 1, 'periods': {1, 2}, 'weeks': {5}, 'location': '实A南-705'}
        """
        pattern = r'\[([^\[\]]*?(?:\[[^\[\]]*?\][^\[\]]*?)*)\]'
        matches = re.findall(pattern, entry_str)
        if len(matches) < 5:
            return None
        teacher = matches[0]
        day = self.parse_day(matches[1])
        periods = self.parse_periods(matches[2])
        weeks = self.parse_weeks(matches[3])
        location = matches[4]
        return {
            'teacher': teacher,
            'day': day,
            'periods': periods,
            'weeks': weeks,
            'location': location
        }

    def parse_schedules(self):
        """
        解析所有课表信息。

        Returns:
            dict: 键为索引，值为包含该索引对应课程所有课表信息的列表。

        Example:
            >>> parser = TimeParser(df, 'time_column')
            >>> parser.parse_schedules()
            {0: [{'teacher': '刘高勇', 'day': 1, 'periods': {1, 2}, 'weeks': {5}, 'location': '实A南-705'}, ...], ...}
        """
        schedules = self.extract_course_time()
        all_schedule_list = {}
        for index, schedule in enumerate(schedules):
            schedule_list = []
            for entry in schedule:
                if entry != '':
                    parsed_entry = self.parse_class_entry(entry)
                    if parsed_entry:
                        schedule_list.append(parsed_entry)
            all_schedule_list[index] = schedule_list
        return all_schedule_list
    

class TimeMatrix:
    """
    用于处理时间矩阵的类，包括冲突检测和时间接近度计算。

    Attributes:
        all_schedule_list_supervisor (dict): 督导的所有课表信息。
        all_schedule_list_task (dict): 任务的所有课表信息。
    """

    def __init__(self, all_schedule_list_supervisor, all_schedule_list_task):
        """
        初始化TimeMatrix对象。

        Args:
            all_schedule_list_supervisor (dict): 督导的所有课表信息。
            all_schedule_list_task (dict): 任务的所有课表信息。
        """
        self.all_schedule_list_supervisor = all_schedule_list_supervisor
        self.all_schedule_list_task = all_schedule_list_task

    def weeks_overlap(self, weeks1, weeks2):
        """
        判断两个周数集合是否有重叠。

        Args:
            weeks1 (set): 第一个周数集合。
            weeks2 (set): 第二个周数集合。

        Returns:
            bool: 如果有重叠返回True，否则返回False。

        Example:
            >>> matrix = TimeMatrix(supervisor_schedules, task_schedules)
            >>> matrix.weeks_overlap({1, 2, 3}, {3, 4, 5})
            True
        """
        return not weeks1.isdisjoint(weeks2)

    def periods_overlap(self, periods1, periods2):
        """
        判断两个节次集合是否有重叠。

        Args:
            periods1 (set): 第一个节次集合。
            periods2 (set): 第二个节次集合。

        Returns:
            bool: 如果有重叠返回True，否则返回False。

        Example:
            >>> matrix = TimeMatrix(supervisor_schedules, task_schedules)
            >>> matrix.periods_overlap({1, 2}, {2, 3})
            True
        """
        return not periods1.isdisjoint(periods2)

    def schedules_conflict(self, schedule1, schedule2):
        """
        判断两个课表列表是否存在时间冲突。

        Args:
            schedule1 (list): 第一个课表列表 (督导的课表)。
            schedule2 (list): 第二个课表列表 (任务的课表)。

        Returns:
            bool: 如果存在冲突返回True，否则返回False。

        Note:
            判断冲突的标准为任务的课表时间完全包含于督导的课表时间内。

        Example:
            >>> matrix = TimeMatrix(supervisor_schedules, task_schedules)
            >>> matrix.schedules_conflict(supervisor_schedule, task_schedule)
            True
        """
        for class_task in schedule2:
            for class_sup in schedule1:
                if class_task['day'] == class_sup['day']:
                    if self.weeks_overlap(class_task['weeks'], class_sup['weeks']):
                        if self.periods_overlap(class_task['periods'], class_sup['periods']):
                            if class_task['weeks'].issubset(class_sup['weeks']) and class_task['periods'].issubset(class_sup['periods']):
                                return True
        return False
    
    def calculate_time_proximity(self, supervisor_schedules, task_schedules):
        """
        计算督导与任务的时间接近度。

        Args:
            supervisor_schedules (list): 督导的课表列表。
            task_schedules (list): 任务的课表列表。

        Returns:
            int: 时间接近度得分。

        Example:
            >>> matrix = TimeMatrix(supervisor_schedules, task_schedules)
            >>> matrix.calculate_time_proximity(supervisor_schedule, task_schedule)
            15
        """
        proximity = 0
        for task_class in task_schedules:
            task_day = task_class['day']
            task_periods = task_class['periods']
            supervisor_periods = set()
            for sup_class in supervisor_schedules:
                if sup_class['day'] == task_day:
                    supervisor_periods.update(sup_class['periods'])
            if supervisor_periods:
                min_diff = min([abs(p - t) for p in supervisor_periods for t in task_periods])
                proximity += (10 - min_diff) if min_diff < 10 else 0
        return proximity
    
    def build_conflict_matrix(self):
        """
        构建冲突矩阵和时间接近度矩阵。

        Returns:
            tuple: 包含冲突矩阵和时间接近度矩阵的元组。
                - conflict (numpy.ndarray): 冲突矩阵，表示每个督导与每个任务是否有时间冲突。
                - proximity (numpy.ndarray): 时间接近度矩阵。

        Example:
            >>> matrix = TimeMatrix(supervisor_schedules, task_schedules)
            >>> conflict, proximity = matrix.build_conflict_matrix()
            >>> print(conflict.shape, proximity.shape)
            (10, 20) (10, 20)
        """
        num_supervisors = len(self.all_schedule_list_supervisor)
        num_tasks = len(self.all_schedule_list_task)
        conflict = np.zeros((num_supervisors, num_tasks), dtype=bool)
        proximity = np.zeros((num_supervisors, num_tasks), dtype=int)
        for i in range(num_supervisors):
            for j in range(num_tasks):
                conflict[i][j] = self.schedules_conflict(self.all_schedule_list_supervisor[i], self.all_schedule_list_task[j])
                if conflict[i][j]:
                    proximity[i][j] = 100  # 高冲突时的惩罚
                else:
                    proximity[i][j] = self.calculate_time_proximity(self.all_schedule_list_supervisor[i], self.all_schedule_list_task[j])
        return conflict, proximity


if __name__ == '__main__':
    from data_preprocess import get_data
    course_file_path = '../test-algorithm/2024春开课任务.xls'
    supervisor_file_path = '../test-algorithm/督导名单.xlsx'
    supervisor_course_data, task_data, _ = get_data(course_file_path, supervisor_file_path)
    time_parser_supervisor = TimeParser(supervisor_course_data, '课表信息')
    time_parser_task = TimeParser(task_data, '课表信息')
    all_schedule_list_supervisor = time_parser_supervisor.parse_schedules()
    all_schedule_list_task = time_parser_task.parse_schedules()
    print(len(all_schedule_list_supervisor))
    print(len(all_schedule_list_task))
    # 构建冲突矩阵和时间接近度矩阵
    time_matrix = TimeMatrix(all_schedule_list_supervisor, all_schedule_list_task)
    conflict, proximity = time_matrix.build_conflict_matrix()
    print(conflict)
    print(proximity)

