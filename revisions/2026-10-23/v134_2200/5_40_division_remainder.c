#include <stdio.h>
int main(void)
{
    int total = 0;
    int group = 0;
    printf("Total and group size: ");
    if (scanf("%d %d", &total, &group) != 2) {
        puts("Input error.");
        return 1;
    } else if (total < 0 || group <= 0) {
        puts("Use a nonnegative total and a positive group size.");
        return 1;
    } else {
        printf("Groups: %d\n", total / group);
        printf("Leftover: %d\n", total % group);
    }
    return 0;
}