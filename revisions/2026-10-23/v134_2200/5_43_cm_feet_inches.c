#include <stdio.h>
#include <math.h>
int main(void)
{
    double centimeters = 0.0;
    double total_inches = 0.0;
    int feet = 0;
    double inches = 0.0;
    printf("Height in centimeters (0..10000): ");
    if (scanf("%lf", &centimeters) != 1 ||
        !isfinite(centimeters) || centimeters < 0.0 || centimeters > 10000.0) {
        puts("Invalid height.");
        return 1;
    } else {
        total_inches = centimeters / 2.54;
        feet = (int)(total_inches / 12.0);
        inches = total_inches - (double)feet * 12.0;
        printf("%d feet %.2f inches\n", feet, inches);
    }
    return 0;
}