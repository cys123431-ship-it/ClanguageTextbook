#include <stdio.h>
int main(void)
{
    int number = 0;
    printf("Two-digit number (10..99): ");
    if (scanf("%d", &number) != 1 ||
        number < 10 || number > 99) {
        puts("Enter a number from 10 to 99.");
        return 1;
    } else {
        printf("Tens: %d\n", number / 10);
        printf("Ones: %d\n", number % 10);
    }
    return 0;
}