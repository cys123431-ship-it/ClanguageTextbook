#include <stdio.h>

int main(void)
{
    int x = 10;
    int y = 20;
    int temporary = 0;
    printf("Before: x=%d y=%d\n", x, y);
    temporary = x;
    x = y;
    y = temporary;
    printf("After: x=%d y=%d\n", x, y);
    return 0;
}
