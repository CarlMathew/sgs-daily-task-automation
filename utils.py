from datetime import datetime, timedelta

class Utils:
    def get_total_workday(days: int) -> int:
        """
            Fetch the total if weekends are not included on the days
        """
        for i in range(days):
            current_day = datetime.now() - timedelta(days=i)
            day_of_week = current_day.strftime("%A")

        if day_of_week in ["Saturday", "Sunday"]:
            days += 1 
        
        return days