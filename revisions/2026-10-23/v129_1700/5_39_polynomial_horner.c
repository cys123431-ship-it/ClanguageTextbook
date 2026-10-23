#include <stdio.h>
#include <math.h>

int main(void)
{
    double x = 0.0;
    printf("x (-1000 to 1000): ");
    if (scanf("%lf", &x) != 1 || !isfinite(x) || x < -1000.0 || x > 1000.0) {
        puts("Invalid x.");
        return 1;
    } else {
        double direct = 2.0 * x * x + 5.0 * x + 9.0;
        double horner = (2.0 * x + 5.0) * x + 9.0;
        printf("Direct: %.2f\n", direct);
        printf("Horner: %.2f\n", horner);
    }
    return 0;
}
