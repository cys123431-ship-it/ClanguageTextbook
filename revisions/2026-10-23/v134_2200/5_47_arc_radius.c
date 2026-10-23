#include <stdio.h>
#include <math.h>
int main(void)
{
    const double pi = 3.141592653589793;
    double arc_km = 0.0;
    double degrees = 0.0;
    printf("Arc length (km) and central angle (degrees): ");
    if (scanf("%lf %lf", &arc_km, &degrees) != 2 ||
        !isfinite(arc_km) || !isfinite(degrees) ||
        arc_km <= 0.0 || arc_km > 1.0e7 ||
        degrees < 0.01 || degrees > 360.0) {
        puts("Invalid length or angle.");
        return 1;
    } else {
        double radius_km = arc_km * 180.0 / (pi * degrees);
        printf("Estimated radius: %.2f km\n", radius_km);
    }
    return 0;
}