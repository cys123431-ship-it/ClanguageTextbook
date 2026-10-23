#include <stdio.h>
#include <math.h>

int main(void)
{
    double meters = 0.0;
    const double meters_per_mile = 1609.344;

    printf("Distance in meters (0 to 10000000): ");
    if (scanf("%lf", &meters) != 1 || !isfinite(meters) || meters < 0.0 || meters > 10000000.0) {
        puts("Invalid distance.");
        return 1;
    } else {
        double kilometers = meters / 1000.0;
        double miles = meters / meters_per_mile;
        printf("%.3f m = %.6f km = %.6f mi\n", meters, kilometers, miles);
    }
    return 0;
}
