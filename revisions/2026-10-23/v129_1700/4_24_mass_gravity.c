#include <stdio.h>
#include <math.h>

int main(void)
{
    double mass_kg = 0.0;
    const double earth_g = 9.81;
    const double moon_g = 1.62;

    printf("Mass in kg (0 to 10000): ");
    if (scanf("%lf", &mass_kg) != 1 || !isfinite(mass_kg) || mass_kg < 0.0 || mass_kg > 10000.0) {
        puts("Invalid mass.");
        return 1;
    } else {
        double earth_newtons = mass_kg * earth_g;
        double moon_newtons = mass_kg * moon_g;
        printf("Mass: %.2f kg\n", mass_kg);
        printf("Earth force: %.2f N\n", earth_newtons);
        printf("Moon force: %.2f N\n", moon_newtons);
    }
    return 0;
}
