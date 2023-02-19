package z_FunctionalLearning.functioanlinterface;

import java.util.function.BiFunction;
import java.util.function.Function;

public class _Function {
    public static void main(String[] args) {
        int increment = increment (1);
        System.out.println (increment);
        System.out.println (incrementByOneFunction.apply (5));
        Function<Integer, Integer> multiplyBy10ThenAdd1 = multipleBy10
                .andThen(incrementByOneFunction);
        System.out.println (multiplyBy10ThenAdd1.apply (5));
        System.out.println (incrementByOneAndMultipleBy10.apply (4, 5));
    }

    static int increment (int number){
        return number+1;
    }

    static BiFunction<Integer, Integer, Integer> incrementByOneAndMultipleBy10 =
            (numberToIncrement, numberToMultiply) -> (numberToIncrement +1)*numberToMultiply;
    static Function<Integer,Integer > incrementByOneFunction = number -> number+1;
    static Function<Integer,Integer > multipleBy10 = number -> number*10;


}
