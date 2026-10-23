#include <stdio.h>
#include <math.h>

int main(void)
{
    double value = 0.0;

    printf("Enter a nonnegative measurement: ");
    if (scanf("%lf", &value) != 1 ||
        !isfinite(value) || value < 0.0 || value > 1000000.0) {
        puts("Invalid measurement.");
        return 1;
    } else {
        printf("Fixed (8 digits): %.8f\n", value);
        printf("Scientific (3 digits): %.3e\n", value);
        printf("General (6 significant digits): %.6g\n", value);
    }
    return 0;
}