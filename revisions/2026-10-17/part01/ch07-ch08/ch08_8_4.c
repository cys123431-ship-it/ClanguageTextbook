#include <stdio.h>

int hidden_counter(void)
{
    static int count = 0;
    count++;
    return count;
}

int add_one(int previous)
{
    return previous + 1;
}

int main(void)
{
    int visible = 0;
    int first = hidden_counter();
    int second = hidden_counter();
    printf("Hidden: %d %d\n", first, second);
    visible = add_one(visible);
    visible = add_one(visible);
    printf("Visible: %d\n", visible);
    return 0;
}
