#include <stdio.h>
#include <math.h>

int main(void)
{
    const double meters_per_mile = 1609.344;
    double miles = 0.0;

    printf("Distance in international miles: ");
    if (scanf("%lf", &miles) != 1 ||
        !isfinite(miles) || miles < 0.0 || miles > 1000000000.0) {
        puts("Invalid distance.");
        return 1;
    } else {
        double meters = miles * meters_per_mile;
        printf("Distance in meters: %.3f\n", meters);
    }
    return 0;
}
