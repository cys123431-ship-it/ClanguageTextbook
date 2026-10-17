#include "inventory.h"
int inventory_total(const int *prices, size_t count) {
    int sum = 0;
    for (size_t i = 0; i < count; ++i) {
        sum += prices[i];
    }
    return sum;
}
