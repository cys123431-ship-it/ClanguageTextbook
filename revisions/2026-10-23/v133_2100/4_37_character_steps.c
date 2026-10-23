#include <stdio.h>

int main(void)
{
    if ('b' != 'a' + 1 || 'c' != 'a' + 2 || 'd' != 'a' + 3) {
        puts("This example needs consecutive a-d codes.");
        return 2;
    } else {
        printf("Letters: %c %c %c\n", 'a' + 1, 'a' + 2, 'a' + 3);
        printf("Codes: %d %d %d\n", 'a' + 1, 'a' + 2, 'a' + 3);
    }
    return 0;
}
