#include <stdio.h>
#include <math.h>

int main(void)
{
    double calculated = 1.0 - 0.9;
    double expected = 0.1;
    double difference = fabs(calculated - expected);
    double tolerance = 1e-12;
    int close_enough = difference <= tolerance;

    printf("허용 오차 안: %s\n", close_enough ? "예" : "아니요");
    printf("절대 오차: %.17g\n", difference);
    return 0;
}
