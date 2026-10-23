#include <stdio.h>
#include <math.h>

int main(void)
{
    const double square_meters_per_pyeong = 3.305785;
    double pyeong = 0.0;

    printf("Area in pyeong (0 to 10000): ");
    if (scanf("%lf", &pyeong) != 1 ||
        !isfinite(pyeong) || pyeong < 0.0 || pyeong > 10000.0) {
        puts("Invalid area.");
        return 1;
    } else {
        double square_meters = pyeong * square_meters_per_pyeong;
        printf("Area: %.3f square meters\n", square_meters);
    }
    return 0;
}
