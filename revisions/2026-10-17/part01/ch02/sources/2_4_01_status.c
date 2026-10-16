#include <stdio.h>

int main(void)
{
    int seats = 12;
    int reserved = 5;
    int available = seats - reserved;

    printf("available=%d\n", available);
    return 0;
}