#include <stdio.h>

static double celsius_to_fahrenheit(double celsius)
{
    return celsius * 9.0 / 5.0 + 32.0;
}

static double fahrenheit_to_celsius(double fahrenheit)
{
    return (fahrenheit - 32.0) * 5.0 / 9.0;
}

int main(void)
{
    const double samples[] = {-40.0, 0.0, 37.0, 100.0};
    const size_t count = sizeof samples / sizeof samples[0];

    for (size_t i = 0; i < count; ++i) {
        double celsius = samples[i];
        double fahrenheit = celsius_to_fahrenheit(celsius);
        printf("%.1f C -> %.1f F\n", celsius, fahrenheit);
    }
    printf("212.0 F -> %.1f C\n", fahrenheit_to_celsius(212.0));
    return 0;
}
