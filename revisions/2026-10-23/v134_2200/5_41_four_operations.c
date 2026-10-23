#include <stdio.h>
#include <math.h>
int main(void)
{
    double left = 0.0;
    double right = 0.0;
    printf("Two real numbers: ");
    if (scanf("%lf %lf", &left, &right) != 2 ||
        !isfinite(left) || !isfinite(right) ||
        left < -1.0e6 || left > 1.0e6 ||
        right < -1.0e6 || right > 1.0e6) {
        puts("Input error.");
        return 1;
    } else {
        printf("Sum: %.2f\n", left + right);
        printf("Difference: %.2f\n", left - right);
        printf("Product: %.2f\n", left * right);
        if (right == 0.0) {
            puts("Quotient: unavailable (division by zero)");
        } else {
            printf("Quotient: %.2f\n", left / right);
        }
    }
    return 0;
}