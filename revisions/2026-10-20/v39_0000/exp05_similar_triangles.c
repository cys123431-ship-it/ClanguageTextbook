#include <stdio.h>
#include <math.h>

int main(void)
{
    double small_height = 0.0;
    double small_base = 0.0;
    double large_base = 0.0;

    if (scanf("%lf %lf %lf", &small_height, &small_base, &large_base) != 3) {
        puts("세 길이를 모두 숫자로 입력하세요.");
        return 1;
    }
    if (!isfinite(small_height) || !isfinite(small_base) ||
        !isfinite(large_base) || small_height <= 0.0 ||
        small_base <= 0.0 || large_base <= 0.0) {
        puts("길이는 모두 양의 유한한 값이어야 합니다.");
        return 1;
    }

    double ratio = small_height / small_base;
    double large_height = ratio * large_base;
    if (!isfinite(large_height)) {
        puts("계산 결과가 너무 큽니다.");
        return 1;
    }
    printf("큰 삼각형의 높이: %.2f\n", large_height);
    return 0;
}
