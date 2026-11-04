CREATE SCHEMA IF NOT EXISTS farm_project;

SET search_path TO farm_project;

-- Очистка (на случай, если таблицы уже были созданы с ошибками)
DROP TABLE IF EXISTS farm_project.laboratory_data CASCADE;
DROP TABLE IF EXISTS farm_project.milk_movement CASCADE;
DROP TABLE IF EXISTS farm_project.events CASCADE;
DROP TABLE IF EXISTS farm_project.milk_yield CASCADE;
DROP TABLE IF EXISTS farm_project.animals CASCADE;
DROP TABLE IF EXISTS farm_project.users CASCADE;
DROP TABLE IF EXISTS farm_project.groups CASCADE;

-- 1. ГРУППЫ
CREATE TABLE groups (
    id BIGINT PRIMARY KEY,
    name VARCHAR(100) NOT NULL
);

INSERT INTO groups VALUES
(1, 'Дойная группа №1'),
(2, 'Дойная группа №2'),
(3, 'Группа молодняка');

-- 2. ПОЛЬЗОВАТЕЛИ
CREATE TABLE users (
    id BIGINT PRIMARY KEY,
    full_name VARCHAR(100) NOT NULL,
    role VARCHAR(30) NOT NULL DEFAULT 'Доярка'
);

ALTER TABLE users
ADD CONSTRAINT ch_user_role
CHECK (role IN ('Доярка', 'Лаборант', 'Бригадир'));

INSERT INTO users VALUES
(1, 'Иванова Анна', 'Доярка'),
(2, 'Иванова Мария', 'Доярка'),
(3, 'Петрова Анна', 'Доярка'),
(4, 'Сидоров Иван', 'Лаборант'),
(5, 'Петров Сергей', 'Лаборант'),
(6, 'Сидорова Мария', 'Бригадир'),
(7, 'Кузнецов Алексей', 'Доярка');

-- 3. ЖИВОТНЫЕ
CREATE TABLE animals (
    id BIGINT PRIMARY KEY,
    tag_number VARCHAR(50) NOT NULL,
    nickname VARCHAR(50),
    group_id BIGINT,
    status VARCHAR(30) NOT NULL DEFAULT 'Дойная',
    date_of_birth DATE
);

ALTER TABLE animals
ADD CONSTRAINT uq_animals_tag_number UNIQUE (tag_number);

ALTER TABLE animals
ADD CONSTRAINT ch_animal_status
CHECK (status IN ('Дойная', 'Сухостой', 'На лечении'));

ALTER TABLE animals
ADD CONSTRAINT fk_animals_group_id
FOREIGN KEY (group_id) REFERENCES groups(id);

INSERT INTO animals VALUES
(1, 'COW-001', 'Бурёнка', 1, 'Дойная', '2022-03-15'),
(2, 'COW-002', 'Звёздочка', 1, 'Дойная', '2021-07-20'),
(3, 'COW-003', 'Милка', 2, 'Сухостой', '2020-11-10'),
(4, 'COW-004', 'Ромашка', 2, 'На лечении', '2023-01-05'),
(5, 'COW-005', 'Ласточка', 3, 'Дойная', '2022-09-12');

-- 4. НАДОИ
CREATE TABLE milk_yield (
    id BIGINT PRIMARY KEY,
    date_yield DATE NOT NULL,
    shift VARCHAR(20) NOT NULL,
    animal_id BIGINT,
    group_id BIGINT,
    volume_liters NUMERIC(8,2) NOT NULL,
    user_id BIGINT
);

ALTER TABLE milk_yield
ADD CONSTRAINT ch_milk_shift CHECK (shift IN ('Утро', 'День', 'Вечер'));

ALTER TABLE milk_yield
ADD CONSTRAINT ch_milk_volume CHECK (volume_liters > 0);

ALTER TABLE milk_yield
ADD CONSTRAINT fk_milk_yield_animal FOREIGN KEY (animal_id) REFERENCES animals(id);

ALTER TABLE milk_yield
ADD CONSTRAINT fk_milk_yield_group FOREIGN KEY (group_id) REFERENCES groups(id);

ALTER TABLE milk_yield
ADD CONSTRAINT fk_milk_yield_user FOREIGN KEY (user_id) REFERENCES users(id);

-- Логика: заполнен либо animal_id, либо group_id
ALTER TABLE milk_yield
ADD CONSTRAINT ch_milk_yield_subject CHECK (
    (animal_id IS NOT NULL AND group_id IS NULL)
    OR
    (animal_id IS NULL AND group_id IS NOT NULL)
);

-- Исправленные данные (все колонки строго по порядку!)
INSERT INTO milk_yield (id, date_yield, shift, animal_id, group_id, volume_liters, user_id) VALUES
(1, '2026-10-01', 'Утро', 1, NULL, 18.50, 1),
(2, '2026-10-01', 'День', 1, NULL, 12.30, 1),
(3, '2026-10-01', 'Вечер', 2, NULL, 16.80, 3),
(4, '2026-10-02', 'Утро', NULL, 1, 125.50, 1),
(5, '2026-10-02', 'День', NULL, 2, 98.40, 2),
(6, '2026-10-03', 'Утро', 5, NULL, 20.10, 1),
-- Данные за 2023 год (для истории, исправлено смещение колонок)
(7, '2023-10-01', 'Утро', 1, NULL, 15.5, 1),
(8, '2023-10-01', 'Вечер', 1, NULL, 14.0, 1),
(9, '2023-10-01', 'Утро', 2, NULL, 12.0, 1),
(10, '2023-10-01', 'Вечер', 2, NULL, 13.2, 1),
(11, '2023-10-02', 'Утро', 1, NULL, 16.0, 1),
(12, '2023-10-02', 'Вечер', 1, NULL, 15.1, 1),
(13, '2023-10-02', 'Утро', 2, NULL, 11.8, 1),
(14, '2023-10-02', 'Вечер', 2, NULL, 12.5, 1),
(15, '2023-10-03', 'Утро', 3, NULL, 10.5, 2),
(16, '2023-10-03', 'Вечер', 3, NULL, 10.0, 2),
(17, '2023-10-03', 'Утро', 5, NULL, 18.2, 2),
(18, '2023-10-03', 'Вечер', 5, NULL, 17.5, 2),
(19, '2023-10-04', 'Утро', 3, NULL, 11.0, 2),
(20, '2023-10-04', 'Вечер', 3, NULL, 9.8, 2),
(21, '2023-10-04', 'Утро', 5, NULL, 19.0, 2),
(22, '2023-10-04', 'Вечер', 5, NULL, 18.1, 2);

-- 5. ЛАБОРАТОРНЫЕ ДАННЫЕ
CREATE TABLE laboratory_data (
    id BIGINT PRIMARY KEY,
    date_analyze DATE NOT NULL,
    animal_id BIGINT,
    yield_id BIGINT,
    fat_percent NUMERIC(5,2),
    protein_percent NUMERIC(5,2),
    acidity NUMERIC(5,2)
);

ALTER TABLE laboratory_data
ADD CONSTRAINT ch_lab_fat CHECK (fat_percent BETWEEN 0 AND 100);

ALTER TABLE laboratory_data
ADD CONSTRAINT ch_lab_protein CHECK (protein_percent BETWEEN 0 AND 100);

ALTER TABLE laboratory_data
ADD CONSTRAINT ch_lab_acidity CHECK (acidity >= 0);

ALTER TABLE laboratory_data
ADD CONSTRAINT fk_lab_animal FOREIGN KEY (animal_id) REFERENCES animals(id);

ALTER TABLE laboratory_data
ADD CONSTRAINT fk_lab_yield FOREIGN KEY (yield_id) REFERENCES milk_yield(id);

ALTER TABLE laboratory_data
ADD CONSTRAINT ch_lab_subject CHECK (
    (animal_id IS NOT NULL AND yield_id IS NULL)
    OR
    (animal_id IS NULL AND yield_id IS NOT NULL)
);

-- Исправлены данные (строгое соблюдение колонок и связей)
INSERT INTO laboratory_data (id, date_analyze, animal_id, yield_id, fat_percent, protein_percent, acidity) VALUES
(1, '2026-10-01', NULL, 1, 3.80, 3.20, 16.50),
(2, '2026-10-01', NULL, 2, 3.70, 3.15, 16.80),
(3, '2026-10-02', 3, NULL, 3.90, 3.30, 16.40),
(4, '2026-10-02', NULL, 4, 3.75, 3.25, 16.60),
(5, '2026-10-03', NULL, 6, 3.85, 3.28, 16.30),
-- Данные за 2023 год (исправлена структура)
(6, '2023-10-01', 1, NULL, 3.80, 3.20, 17.50),
(7, '2023-10-01', 2, NULL, 4.10, 3.30, 18.00),
(8, '2023-10-02', 1, NULL, 3.70, 3.10, 17.20),
(9, '2023-10-02', 2, NULL, 4.00, 3.20, 17.80),
(10, '2023-10-03', 3, NULL, 3.90, 3.40, 16.50), 
(11, '2023-10-04', 5, NULL, 4.20, 3.50, 16.80);

-- 6. СОБЫТИЯ
CREATE TABLE events (
    id BIGINT PRIMARY KEY,
    animal_id BIGINT NOT NULL,
    event_date DATE NOT NULL,
    event_type VARCHAR(30) NOT NULL,
    description VARCHAR(500),
    withdrawal_date DATE
);

ALTER TABLE events
ADD CONSTRAINT ch_event_type CHECK (
    event_type IN ('Отёл', 'Осеменение', 'Лечение', 'Перевод')
);

ALTER TABLE events
ADD CONSTRAINT fk_events_animal FOREIGN KEY (animal_id) REFERENCES animals(id);

ALTER TABLE events
ADD CONSTRAINT ch_withdrawal_date CHECK (
    withdrawal_date IS NULL OR withdrawal_date >= event_date
);

INSERT INTO events VALUES
(1, 1, '2026-02-10', 'Отёл', 'Рождение одного телёнка', NULL),
(2, 2, '2026-03-15', 'Осеменение', 'Плановое осеменение', NULL),
(3, 4, '2026-09-28', 'Лечение', 'Ветеринарный осмотр', '2026-10-05'),
(4, 5, '2026-04-20', 'Перевод', 'Перевод в дойную группу', NULL),
(5, 3, '2026-09-15', 'Осеменение', 'Повторное осеменение', NULL);

-- 7. ДВИЖЕНИЕ МОЛОКА
CREATE TABLE milk_movement (
    id BIGINT PRIMARY KEY,
    date_movement DATE NOT NULL,
    movement_type VARCHAR(20) NOT NULL,
    direction VARCHAR(50),
    volume_liters NUMERIC(8,2) NOT NULL,
    comment VARCHAR(500),
    user_id BIGINT
);

ALTER TABLE milk_movement
ADD CONSTRAINT ch_movement_type CHECK (movement_type IN ('Приход', 'Расход'));

-- Разрешенные направления расхода
ALTER TABLE milk_movement
ADD CONSTRAINT ch_movement_direction CHECK (
    direction IS NULL
    OR direction IN ('Отгрузка на завод', 'Выпойка телят', 'Столовая')
);

ALTER TABLE milk_movement
ADD CONSTRAINT ch_movement_volume CHECK (volume_liters > 0);

ALTER TABLE milk_movement
ADD CONSTRAINT fk_milk_movement_user FOREIGN KEY (user_id) REFERENCES users(id);

-- Исправлено: 'Реализация' заменена на 'Отгрузка на завод', чтобы пройти CHECK
INSERT INTO milk_movement VALUES
(1, '2026-10-01', 'Приход', NULL, 350.00, 'Приём молока после утреннего доения', 1),
(2, '2026-10-01', 'Расход', 'Отгрузка на завод', 280.00, 'Плановая отгрузка', 3),
(3, '2026-10-01', 'Расход', 'Выпойка телят', 25.00, 'Молоко для молодняка', 1),
(4, '2026-10-02', 'Приход', NULL, 390.50, 'Поступление молока за смену', 4),
(5, '2026-10-02', 'Расход', 'Столовая', 10.00, 'Выдача молока в столовую', 3),
(6, '2026-10-03', 'Расход', 'Отгрузка на завод', 310.00, 'Утренняя отгрузка', 3),
(7, '2023-10-01', 'Расход', 'Отгрузка на завод', 54.70, 'Реализация за день', 1), 
(8, '2023-10-02', 'Расход', 'Отгрузка на завод', 50.00, 'Реализация за день', 1),
(9, '2023-10-02', 'Расход', 'Выпойка телят', 5.40, 'Молоко для молодняка', 1),
(10, '2023-10-03', 'Расход', 'Отгрузка на завод', 56.20, 'Реализация за день', 2),
(11, '2023-10-04', 'Расход', 'Выпойка телят', 6.00, 'Молоко для молодняка', 2);

DELETE FROM milk_movement WHERE date_movement >= '2023-10-01';

INSERT INTO milk_movement (id, date_movement, movement_type, direction, volume_liters, user_id) VALUES 
-- 2 октября отгрузили все, что надоили 1-го и 2-го
(12, '2023-10-02', 'Расход', 'Отгрузка на завод', 100.00, 1),
-- 2 октября выпоили телят
(13, '2023-10-02', 'Расход', 'Выпойка телят', 5.40, 1),
-- 4 октября отгрузили остатки
(14, '2023-10-04', 'Расход', 'Отгрузка на завод', 50.00, 1);
