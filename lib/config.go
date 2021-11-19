package lib

import "github.com/spf13/viper"

type Config struct {
	Server   string `mapstructure:"SERVER"`
	Port     int    `mapstructure:"PORT"`
	Login    string `mapstructure:"LOGIN"`
	Password string `mapstructure:"PASSWORD"`
	From     string `mapstructure:"FROM"`
}

// LoadConfig reads configuration from file or environment variables.
func LoadConfig(path string) (config Config, err error) {
	viper.AddConfigPath(path)
	viper.SetConfigName("smail")
	viper.SetConfigType("env")

	viper.AutomaticEnv()

	err = viper.ReadInConfig()
	if err != nil {
		return
	}

	err = viper.Unmarshal(&config)
	return
}
