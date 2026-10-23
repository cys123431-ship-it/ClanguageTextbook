#include <stdio.h>

int main(void)
{
    double first = 3.32e-3;
    double second = 9.76e-5;
    double total = first + second;
    printf("First: %.7f\n", first);
    printf("Second: %.7f\n", second);
    printf("Sum: %.7f\n", total);
    return 0;
}
