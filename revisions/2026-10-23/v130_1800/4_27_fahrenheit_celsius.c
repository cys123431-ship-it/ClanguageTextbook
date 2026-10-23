#include <stdio.h>
#include <math.h>

int main(void)
{
    double fahrenheit = 0.0;

    printf("Temperature in Fahrenheit: ");
    if (scanf("%lf", &fahrenheit) != 1 ||
        !isfinite(fahrenheit) ||
        fahrenheit < -459.67 || fahrenheit > 1000000.0) {
        puts("Invalid temperature.");
        return 1;
    } else {
        double celsius = (5.0 / 9.0) * (fahrenheit - 32.0);
        printf("Temperature in Celsius: %.2f\n", celsius);
    }
    return 0;
}
