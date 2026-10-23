#include <stdio.h>
#include <math.h>
int main(void)
{
    const double pi = 3.141592653589793;
    double radius = 0.0;
    printf("Radius (0..1000): ");
    if (scanf("%lf", &radius) != 1 ||
        !isfinite(radius) || radius < 0.0 || radius > 1000.0) {
        puts("Invalid radius.");
        return 1;
    } else {
        double area = 4.0 * pi * radius * radius;
        double volume = (4.0 / 3.0) * pi * radius * radius * radius;
        printf("Surface area: %.2f\n", area);
        printf("Volume: %.2f\n", volume);
    }
    return 0;
}