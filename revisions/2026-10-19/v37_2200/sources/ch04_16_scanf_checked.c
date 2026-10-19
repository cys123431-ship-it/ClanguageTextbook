#include <stdio.h>

int main(void)
{
    int quantity = 0;
    printf("quantity (1-10): ");

    if (scanf("%d", &quantity) != 1) {
        puts("input error");
        return 1;
    }
    if (quantity < 1 || quantity > 10) {
        puts("out of range");
        return 1;
    }
    printf("accepted: %d\n", quantity);
    return 0;
}
