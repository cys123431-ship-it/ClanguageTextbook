#include <stdio.h>
int main(void)
{
    int a = 0;
    int b = 0;
    int c = 0;
    int largest = 0;
    printf("Three integers: ");
    if (scanf("%d %d %d", &a, &b, &c) != 3) {
        puts("Input error.");
        return 1;
    } else {
        largest = (a > b) ? a : b;
        largest = (largest > c) ? largest : c;
        printf("Largest: %d\n", largest);
    }
    return 0;
}