package z_FunctionalLearning.imperative;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

import static z_FunctionalLearning.imperative.Main.Gender.*;

public class Main {
    public static void main(String[] args) {
        List<Person> people = List.of (
                new Person ("One", MALE),
                new Person ("Two", FEMALE),
                new Person ("Three", MALE),
                new Person ("Four", FEMALE),
                new Person ("Five", MALE),
                new Person ("Six", FEMALE)
                );
        // IMPERATIVE APPROACH
        System.out.println ("IMPERATIVE APPROACH");
        List<Person> females = new ArrayList<> ();
        for (Person person : people){
            if (FEMALE.equals (person.gender)){
                females.add (person);
            }
        }
        for (Person female :females){
            System.out.println (female);
        }

        // DECLARATIVE METHOD
        System.out.println ("DECLARATIVE METHOD");
        people.stream ()
                .filter (person -> FEMALE.equals (person.gender))
                .collect (Collectors.toList ())
                .forEach (System.out::println);
    }

    static class Person{
        private final String name;
        private final Gender gender;

        public Person(String name, Gender gender) {
            this.name = name;
            this.gender = gender;
        }

        @Override
        public String toString() {
            return "Person{" +
                    "name='" + name + '\'' +
                    ", gender=" + gender +
                    '}';
        }
    }
    enum Gender{
        MALE, FEMALE
    }
}
