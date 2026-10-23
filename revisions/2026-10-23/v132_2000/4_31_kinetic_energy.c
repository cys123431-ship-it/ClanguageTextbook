#include <stdio.h>
#include <math.h>

int main(void)
{
    double mass_kg = 0.0;
    double speed_mps = 0.0;

    printf("Mass (kg) and speed (m/s): ");
    if (scanf("%lf %lf", &mass_kg, &speed_mps) != 2 ||
        !isfinite(mass_kg) || !isfinite(speed_mps) ||
        mass_kg < 0.0 || mass_kg > 10000.0 ||
        speed_mps < 0.0 || speed_mps > 1000.0) {
        puts("Invalid mass or speed.");
        return 1;
    } else {
        double joules = 0.5 * mass_kg * speed_mps * speed_mps;
        printf("Kinetic energy: %.2f J\n", joules);
    }
    return 0;
}
