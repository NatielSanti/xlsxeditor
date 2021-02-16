package ru.natiel.xlsxeditor.config;

import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.boot.jdbc.DataSourceBuilder;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.jdbc.core.JdbcTemplate;

import javax.sql.DataSource;

@Configuration
public class WebConfig {

    @Bean(name = "dbPrd")
    @ConfigurationProperties(prefix = "app.prod.datasource")
    public DataSource dataSourcePrd() {
        return DataSourceBuilder.create().build();
    }

    @Bean(name = "jdbcTemplatePrd")
    public JdbcTemplate jdbcTemplatePrd(@Qualifier("dbPrd") DataSource ds) {
        return new JdbcTemplate(ds);
    }

    @Bean(name = "dbStg")
    @ConfigurationProperties(prefix = "app.stage.datasource")
    public DataSource dataSourceStg() {
        return  DataSourceBuilder.create().build();
    }

    @Bean(name = "jdbcTemplateStg")
    public JdbcTemplate jdbcTemplateStg(@Qualifier("dbStg") DataSource ds) {
        return new JdbcTemplate(ds);
    }
}
