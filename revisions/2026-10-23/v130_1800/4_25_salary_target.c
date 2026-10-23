#include <stdio.h>
#include <math.h>

int main(void)
{
    const double target_manwon = 100000.0;
    double salary_manwon = 0.0;

    printf("Annual salary in 10000 KRW: ");
    if (scanf("%lf", &salary_manwon) != 1 ||
        !isfinite(salary_manwon) ||
        salary_manwon <= 0.0 || salary_manwon > 1000000.0) {
        puts("Invalid annual salary.");
        return 1;
    } else {
        double years = target_manwon / salary_manwon;
        printf("Years to reach 1 billion KRW: %.2f\n", years);
    }
    return 0;
}
