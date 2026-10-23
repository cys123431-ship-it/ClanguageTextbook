#include <stdio.h>

#define DAYS_IN_WEEK 7

int main(void)
{
    const double won_per_hour = 500.0;
    int hours_per_day = 0;
    int weeks = 0;

    printf("Daily hours and weeks: ");
    if (scanf("%d %d", &hours_per_day, &weeks) != 2 ||
        hours_per_day < 0 || hours_per_day > 24 ||
        weeks < 1 || weeks > 52) {
        puts("Invalid schedule.");
        return 1;
    } else {
        int total_hours = hours_per_day * DAYS_IN_WEEK * weeks;
        double total_won = (double)total_hours * won_per_hour;
        printf("Total hours: %d\n", total_hours);
        printf("Estimated cost: %.0f KRW\n", total_won);
    }
    return 0;
}