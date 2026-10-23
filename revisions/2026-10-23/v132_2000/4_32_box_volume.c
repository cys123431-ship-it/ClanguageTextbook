#include <stdio.h>
#include <math.h>

int main(void)
{
    double length = 0.0;
    double width = 0.0;
    double height = 0.0;

    printf("Box length, width, height (cm): ");
    if (scanf("%lf %lf %lf", &length, &width, &height) != 3 ||
        !isfinite(length) || !isfinite(width) || !isfinite(height) ||
        length <= 0.0 || width <= 0.0 || height <= 0.0 ||
        length > 1000.0 || width > 1000.0 || height > 1000.0) {
        puts("Invalid dimensions.");
        return 1;
    } else {
        double volume = length * width * height;
        printf("Volume: %.2f cubic cm\n", volume);
    }
    return 0;
}
