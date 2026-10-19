#include <stdio.h>
#include <stddef.h>

static size_t lower_bound_int(const int *a, size_t n, int key)
{
    size_t lo = 0, hi = n;
    while (lo < hi) {
        size_t mid = lo + (hi - lo) / 2;
        if (a[mid] < key) lo = mid + 1;
        else hi = mid;
    }
    return lo;
}

static size_t upper_bound_int(const int *a, size_t n, int key)
{
    size_t lo = 0, hi = n;
    while (lo < hi) {
        size_t mid = lo + (hi - lo) / 2;
        if (a[mid] <= key) lo = mid + 1;
        else hi = mid;
    }
    return lo;
}

int main(void)
{
    const int values[] = {1, 3, 3, 3, 5, 7, 7, 9};
    const int queries[] = {3, 7, 0};
    size_t n = sizeof values / sizeof values[0];
    for (size_t i = 0; i < sizeof queries / sizeof queries[0]; ++i) {
        size_t first = lower_bound_int(values, n, queries[i]);
        size_t after = upper_bound_int(values, n, queries[i]);
        printf("값 %d: [%zu, %zu), 개수 %zu\n",
               queries[i], first, after, after - first);
    }
    return 0;
}
