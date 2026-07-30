CREATE TABLE `bookings` (
  `id` int(11) DEFAULT NULL,
  `lastname` varchar(255) DEFAULT NULL,
  `middlename` varchar(255) DEFAULT NULL,
  `firstname` varchar(255) DEFAULT NULL,
  `fullname` varchar(1024) DEFAULT NULL,
  `sex` varchar(1) DEFAULT NULL,
  `dob` varchar(10) DEFAULT NULL,
  `booked_on` varchar(10) DEFAULT NULL,
  `booked_at` varchar(10) DEFAULT NULL,
  `charges` varchar(2048) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_uca1400_ai_ci;

