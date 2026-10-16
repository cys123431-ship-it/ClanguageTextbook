#include <stdio.h>

int main(void)
{
    double first = 0.1;
    double second = 0.2;
    double total = first + second;

    printf("보기 편한 출력: %.2f\n", total);
    printf("자세한 출력: %.17g\n", total);
    return 0;
}
