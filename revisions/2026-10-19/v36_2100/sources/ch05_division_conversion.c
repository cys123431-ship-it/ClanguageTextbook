#include <stdio.h>

int main(void)
{
    int a = 7;
    int b = 2;

    double late = a / b;
    double early = (double)a / b;
    double literal = 7.0 / 2;

    printf("late: %.1f\n", late);
    printf("early: %.1f\n", early);
    printf("literal: %.1f\n", literal);

    return 0;
}
