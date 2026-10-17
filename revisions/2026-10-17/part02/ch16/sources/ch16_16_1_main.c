#include <stdio.h>
int main(void) {
    int price = 1000;
    int *place = &price;
    printf("before=%d\n", price);
    *place += 250;
    printf("after=%d same=%d\n", price, place == &price);
    return 0;
}
