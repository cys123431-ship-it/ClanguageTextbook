#include <stdio.h>
#include <math.h>

int main(void)
{
    double base = 0.0;
    double height = 0.0;

    printf("Triangle base: ");
    if (scanf("%lf", &base) != 1) {
        puts("Invalid base.");
        return 1;
    } else {
        printf("Triangle height: ");
        if (scanf("%lf", &height) != 1 ||
            !isfinite(base) || !isfinite(height) ||
            base < 0.0 || height < 0.0 ||
            base > 1000000.0 || height > 1000000.0) {
            puts("Invalid dimensions.");
            return 1;
        } else {
            double area = 0.5 * base * height;
            printf("Triangle area: %.2f\n", area);
        }
    }
    return 0;
}
