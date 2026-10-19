#include <stdio.h>

int main(void)
{
    double width = 0.0;
    double height = 0.0;

    printf("가로 세로 입력: ");
    if (scanf("%lf %lf", &width, &height) != 2) {
        puts("입력 오류");
        return 1;
    }
    if (width <= 0.0 || height <= 0.0) {
        puts("길이는 양수여야 합니다");
        return 1;
    }

    double area = width * height;
    double perimeter = 2.0 * (width + height);
    printf("넓이: %.2f\n", area);
    printf("둘레: %.2f\n", perimeter);
    return 0;
}
