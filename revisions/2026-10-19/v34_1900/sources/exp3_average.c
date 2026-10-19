#include <stdio.h>

int main(void)
{
    double first = 0.0;
    double second = 0.0;
    double third = 0.0;

    printf("세 점수 입력: ");
    if (scanf("%lf %lf %lf", &first, &second, &third) != 3) {
        puts("입력 오류");
        return 1;
    }

    double sum = first + second + third;
    double average = sum / 3.0;
    printf("평균: %.2f\n", average);
    return 0;
}
